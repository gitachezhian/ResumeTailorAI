"""Claude calls: tailor resume, write cover letter, shrink-to-fit."""
from __future__ import annotations

import json
import threading

import anthropic

MODEL = "claude-opus-5-5"
BETAS = ["server-side-fallback-2026-07-01"]

RESUME_SCHEMA = {
    "type": "object",
    "properties": {
        "company": {"type": "string", "description": "Hiring company name from the JD, or empty"},
        "role_title": {"type": "string", "description": "Job title exactly as written in the JD"},
        "industry": {"type": "string", "description": "Industry / domain, e.g. 'Fintech - B2B payments SaaS'"},
        "industry_notes": {"type": "string", "description": "1-2 sentences: what this industry/employer values and how the resume was angled"},
        "keywords": {"type": "array", "items": {"type": "string"}, "description": "8-20 most important JD keywords, priority order"},
        "paragraphs": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "src": {"type": "integer"},
                    "text": {"type": "string"},
                },
                "required": ["src", "text"],
                "additionalProperties": False,
            },
        },
        "gaps": {
            "type": "array",
            "description": "What the JD asks for that the ORIGINAL resume does not show, most severe first",
            "items": {
                "type": "object",
                "properties": {
                    "severity": {"type": "string", "enum": ["high", "medium", "low"]},
                    "title": {"type": "string", "description": "Short name of the gap, 2-6 words"},
                    "detail": {"type": "string", "description": "One sentence: what the JD wants and what the resume is missing"},
                    "suggestion": {"type": "string", "description": "One sentence: how to close or address it (course, project, how to talk about it)"},
                },
                "required": ["severity", "title", "detail", "suggestion"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["company", "role_title", "industry", "industry_notes", "keywords", "paragraphs", "gaps"],
    "additionalProperties": False,
}

FIX_SCHEMA = {
    "type": "object",
    "properties": {
        "edits": {
            "type": "array",
            "description": "One entry per paragraph you changed",
            "items": {
                "type": "object",
                "properties": {
                    "item": {"type": "integer", "description": "Index of the paragraph in the list"},
                    "text": {"type": "string", "description": "The full new text of that paragraph, or empty to delete it"},
                },
                "required": ["item", "text"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["edits"],
    "additionalProperties": False,
}

FIX_COVER_SCHEMA = {
    "type": "object",
    "properties": {"letter": {"type": "string", "description": "The full revised cover letter"}},
    "required": ["letter"],
    "additionalProperties": False,
}

SHRINK_SCHEMA = {
    "type": "object",
    "properties": {"paragraphs": RESUME_SCHEMA["properties"]["paragraphs"]},
    "required": ["paragraphs"],
    "additionalProperties": False,
}

WRITING_RULES = """WRITING RULES (follow strictly; the most important rule: write like an experienced professional communicating clearly, not like an AI trying to sound impressive)

Style
- Simple, clear, everyday English. Prefer common words over complicated vocabulary.
- Short, concise sentences. Get straight to the point; no filler words or unnecessary explanations.
- Natural, human tone: professional but conversational. It should sound like something a real professional would naturally say. Avoid robotic or AI-sounding language.
- No fancy words, jargon or corporate buzzwords unless they are relevant keywords from the job description.
- No em dashes.

Resume and application writing
- Strong, clear action verbs. Keep bullet points short and focused.
- Show skills, achievements and impact without exaggerating.
- Use job-description keywords naturally. Do not force a keyword where it does not fit.
- Avoid generic phrases and overused expressions. Do not make it sound overly polished, artificial or exaggerated.

Sentence structure
- Prefer short sentences over long, complex ones. One idea per sentence.
- Avoid unnecessary adjectives and adverbs.
- Use active voice wherever possible.

General
- When in doubt, choose simpler words, shorter sentences and a more natural tone."""

RESUME_SYSTEM = """You are an expert resume writer and ATS specialist. You tailor an existing resume to one job description by rewriting the text of specific paragraphs. The document's formatting is handled by software: you only return text.

INPUT: the job description, and the resume as numbered paragraphs:
  [id] (kind, LOCKED|EDIT) text
Kinds: contact, heading, entry (employer/role/date lines), bullet, text (profile, skills lines, etc.).
`**x**` marks bold; ` | ` inside a line marks a tab.

OUTPUT `paragraphs`: the EDIT paragraphs only, in document order, as {"src": id, "text": "..."}.
- Each EDIT paragraph may appear zero times (drop it: irrelevant bullet), once (rewrite it), or several times in a row (extra bullets for the same role, using the same src). Never output LOCKED ids. Never move a paragraph to a different role or section.
- Keep the paragraph's kind: a bullet stays one sentence-style bullet (no leading bullet symbol); a skills line keeps its exact bold label prefix (e.g. "**Tools:** ..." stays "**Tools:** ...", copy the label and colon placement exactly) and keeps the same separator style (e.g. " • ").
- No line breaks inside a paragraph text. No tabs.

STEP 1 - read the industry. Identify the industry, product type and seniority (e.g. healthtech B2C mobile, enterprise SaaS for construction, XR training simulation, gaming, govtech). Note the vocabulary, tools, metrics and values that industry rewards.

STEP 2 - rewrite.
- MATCH THE JD: rewrite the EDIT paragraphs so the resume matches the job description as closely as possible. You may add any skills, tools, responsibilities, achievements and metrics the JD asks for, even if they are not in the original resume. Fit them naturally into the existing roles.
- Skills lines: include every skill and tool the JD asks for, front-loaded; drop clearly irrelevant ones.
- Profile: 2-4 sentences aimed squarely at this role and industry; mention the target domain naturally.
- Bullets: lead with the most JD-relevant achievements for each role; strong verbs; concrete outcomes and the metrics already in the resume. Most relevant role gets the most bullets; older/less relevant roles get fewer.
- Match the resume's existing spelling variety (e.g. Australian/British English: specialise, optimise).
- Follow the WRITING RULES below for every word you write.

STEP 3 - bold keywords and metrics. In every bullet and the profile, bold 1-3 items with **double asterisks**: keywords taken from the job description, and important metrics (e.g. **80%**, **40+ students**, **1,000+ icons**). Always bold a metric when a bullet has one. Bold short phrases (1-4 words), never whole clauses. In skills lines keep the bold label and additionally bold the few skills that are direct JD matches. Do not overdo it: roughly 10-15% of words bold at most.

STEP 4 - gaps. In `gaps`, list what the JD asks for that the ORIGINAL resume (as given to you, before your rewrite) does not clearly show: missing skills, tools, domain or industry experience, seniority, years, qualifications, certifications, location or work rights. Rate each one:
- high: a stated must-have or core requirement the resume does not show at all.
- medium: an important or "nice to have" requirement that is missing, or a core one that is only weakly shown.
- low: a minor preference, or something easily explained in an interview.
Order most severe first. Be honest and specific. Do not list gaps the resume clearly covers. 3-10 gaps is typical; an empty list is fine if it is a strong match.

LENGTH IS A HARD LIMIT: the final document must fit on 2 pages. Keep the total characters of your EDIT text at or below the budget given. When in doubt, cut weaker bullets rather than writing long ones.

""" + WRITING_RULES

COVER_SYSTEM = """You write tailored cover letters that sound like a real, experienced professional wrote them.

Rules:
- Match the job description as closely as possible. You may add experience, skills, tools, achievements and metrics the JD asks for, even if they are not in the resume. Fit them naturally into the candidate's existing roles.
- Identify the industry and what this employer cares about; connect 2-3 strong, relevant achievements to the role's key needs, using the JD's terminology.
- Structure: greeting ("Dear Hiring Manager," unless a name is in the JD), a hook opening that names the role and company, 2 body paragraphs (optionally a short list of 3 bullet points of matched strengths), a confident closing paragraph, sign-off with the candidate's name.
- 250-350 words. Specific and confident, never gushing. No clichés like "I am writing to express my interest".
- Match the resume's spelling variety (e.g. Australian English).
- Formatting: plain text with blank lines between paragraphs. You may use "- " bullet lines. Bold (**like this**) the key job-description keywords and important metrics, a handful in total. No headings, no address block, no date line, no placeholders like [Company].
- Put the letter in the `letter` field. Also return the hiring `company` name and the exact `role_title` as written in the job ad.

""" + WRITING_RULES


class Cancelled(Exception):
    """The user pressed Stop."""


class CancelToken:
    """Set per job; Stop flips it and closes the open stream so we stop paying for tokens."""

    def __init__(self):
        self.cancelled = False
        self._streams = []
        self._lock = threading.Lock()

    def attach(self, stream):
        with self._lock:
            self._streams.append(stream)
            if self.cancelled:
                self._close(stream)

    def cancel(self):
        with self._lock:
            self.cancelled = True
            for s in self._streams:
                self._close(s)

    @staticmethod
    def _close(stream):
        try:
            stream.close()
        except Exception:
            pass


_local = threading.local()


def use_token(token: CancelToken | None):
    """Bind a cancel token to the current thread."""
    _local.token = token


def _client(api_key: str | None):
    return anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()


def _call(client, system: str, user: str, schema: dict | None, effort: str = "high", max_tokens: int = 32000) -> str:
    kwargs = dict(
        model=MODEL,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
        betas=BETAS,
        fallbacks="default",
    )
    oc = {"effort": effort}
    if schema:
        oc["format"] = {"type": "json_schema", "schema": schema}
    kwargs["output_config"] = oc
    token = getattr(_local, "token", None)
    if token and token.cancelled:
        raise Cancelled()
    try:
        with client.beta.messages.stream(**kwargs) as stream:
            if token:
                token.attach(stream)
            msg = stream.get_final_message()
    except Exception:
        if token and token.cancelled:
            raise Cancelled()
        raise
    if token and token.cancelled:
        raise Cancelled()
    if msg.stop_reason == "refusal":
        raise RuntimeError("Claude declined this request. Try rephrasing the job description.")
    if msg.stop_reason == "max_tokens":
        raise RuntimeError("The response was cut off (too long). Please try again.")
    text = "".join(b.text for b in msg.content if b.type == "text")
    return text


def tailor_resume(api_key, jd: str, outline: str, char_budget: int, template_name: str) -> dict:
    client = _client(api_key)
    user = (
        f"<job_description>\n{jd.strip()}\n</job_description>\n\n"
        f"<resume template=\"{template_name}\">\n{outline}\n</resume>\n\n"
        f"Character budget for all EDIT text combined (excluding ** markers): about {char_budget} characters maximum.\n"
        "Return the JSON."
    )
    return json.loads(_call(client, RESUME_SYSTEM, user, RESUME_SCHEMA))


def shrink_resume(api_key, jd: str, outline: str, current_items: list[dict], over_by_lines: int, char_budget: int) -> list[dict]:
    client = _client(api_key)
    user = (
        f"<job_description>\n{jd.strip()}\n</job_description>\n\n"
        f"<resume_template>\n{outline}\n</resume_template>\n\n"
        f"<current_tailored_paragraphs>\n{json.dumps(current_items, ensure_ascii=False, indent=1)}\n</current_tailored_paragraphs>\n\n"
        f"The tailored resume is too long: about {over_by_lines} lines over the 2-page limit. "
        f"Shorten it to roughly {char_budget} characters of EDIT text total: cut the least relevant bullets first "
        "(especially in older roles), then tighten wording. Keep all bolding rules and the most JD-relevant content. "
        "Return the full corrected `paragraphs` list (same format)."
    )
    return json.loads(_call(client, RESUME_SYSTEM, user, SHRINK_SCHEMA, effort="medium"))["paragraphs"]


COVER_SCHEMA = {
    "type": "object",
    "properties": {
        "company": {"type": "string", "description": "Hiring company name from the JD, or empty"},
        "role_title": {"type": "string", "description": "Job title exactly as written in the JD"},
        "letter": {"type": "string", "description": "The full cover letter text"},
    },
    "required": ["company", "role_title", "letter"],
    "additionalProperties": False,
}


def cover_letter(api_key, jd: str, resume_text: str) -> dict:
    client = _client(api_key)
    user = (
        f"<job_description>\n{jd.strip()}\n</job_description>\n\n"
        f"<candidate_resume>\n{resume_text}\n</candidate_resume>\n\n"
        "Write the cover letter. Return the JSON."
    )
    return json.loads(_call(client, COVER_SYSTEM, user, COVER_SCHEMA, effort="high", max_tokens=16000))


FIX_SYSTEM = """You revise a tailored resume based on the candidate's own review comments.

INPUT: the job description, the resume's editable paragraphs as a numbered list ([index] text), and a list of comments. Each comment quotes the exact text it is about, names the paragraph index, and says what to change.

Rules:
- Address every comment. Change only what the comments ask for, plus anything a fix makes necessary.
- Return in `edits` only the paragraphs you changed, each with its index and full new text. Empty text deletes the paragraph.
- Keep the paragraph's kind and format: bullets stay single bullets with no leading symbol; skills lines keep their exact bold label prefix and separator style. No line breaks or tabs.
- Keep the bolding style: **double asterisks** around job-description keywords and important metrics.
- Keep a similar length unless a comment asks otherwise. The resume must still fit on 2 pages.

""" + WRITING_RULES

FIX_COVER_SYSTEM = """You revise a cover letter based on the candidate's own review comments.

INPUT: the job description, the current letter, and a list of comments. Each comment quotes the exact text it is about and says what to change.

Rules:
- Address every comment. Change only what the comments ask for, plus anything a fix makes necessary. Keep everything else word for word.
- Keep the same formatting: plain text, blank lines between paragraphs, optional "- " bullet lines, **bold** for key job-description keywords and metrics.
- Return the full revised letter in `letter`.

""" + WRITING_RULES


def fix_resume(api_key, jd: str, items: list[dict], comments: list[dict]) -> list[dict]:
    """items: [{"src", "text"}]; comments: [{"item", "quote", "comment"}]. Returns [{"item", "text"}]."""
    client = _client(api_key)
    listing = "\n".join(f"[{i}] {it['text']}" for i, it in enumerate(items))
    notes = "\n".join(
        f"- Paragraph [{c['item']}], about \"{c['quote']}\": {c['comment']}" for c in comments
    )
    user = (
        f"<job_description>\n{jd.strip()}\n</job_description>\n\n"
        f"<resume_paragraphs>\n{listing}\n</resume_paragraphs>\n\n"
        f"<comments>\n{notes}\n</comments>\n\n"
        "Apply all the comments. Return the JSON."
    )
    return json.loads(_call(client, FIX_SYSTEM, user, FIX_SCHEMA, effort="medium", max_tokens=16000))["edits"]


def fix_cover(api_key, jd: str, letter: str, comments: list[dict]) -> str:
    """comments: [{"quote", "comment"}]. Returns the revised letter."""
    client = _client(api_key)
    notes = "\n".join(f"- About \"{c['quote']}\": {c['comment']}" for c in comments)
    user = (
        f"<job_description>\n{jd.strip()}\n</job_description>\n\n"
        f"<letter>\n{letter}\n</letter>\n\n"
        f"<comments>\n{notes}\n</comments>\n\n"
        "Apply all the comments. Return the JSON."
    )
    return json.loads(_call(client, FIX_COVER_SYSTEM, user, FIX_COVER_SCHEMA, effort="medium", max_tokens=16000))["letter"]


def check_key(api_key) -> None:
    """Cheap call to validate a key."""
    client = _client(api_key)
    client.models.retrieve(MODEL)

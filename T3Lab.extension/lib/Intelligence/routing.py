# -*- coding: utf-8 -*-
"""
Routing decisions for the T3Lab Assistant.

The pure decision logic lifted out of script.py's _route_input(), which was
~880 lines of routing rules interleaved with WPF dispatching and therefore
impossible to test. Everything here takes plain values and returns plain
values: no Revit API, no WPF, no LLM calls. script.py keeps the I/O.

Pure Python, CPython-3 testable — same contract as agents/dispatcher.py.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Routing"

import re

from Intelligence.knowledge import vi_text


# ─── Continuation (answering the assistant's own question) ───────────────────
# When the assistant asks a clarifying question, the user's short reply ("1",
# "toàn bộ project", "ok") is the ANSWER to it, and the turn must continue with
# the same specialist/skill and full history instead of being routed from
# scratch — where it matches nothing and the model loses the thread.
#
# The old test was `'?' in last_bot[-250:] and len(raw.strip()) <= 80`. Both
# halves were too loose: a question mark ANYWHERE in the last 250 characters
# counted (including one inside a mid-message aside), and 80 characters is
# long enough for a complete new command. "tô vàng sàn" typed right after a
# clarifying question satisfied both, so a brand-new command rode the previous
# turn's decision and replayed the finished task. Only the keyword escape
# hatch caught it — one check standing alone. Now the shape of the reply has
# to look like an answer too.

# Longest reply still treated as an answer rather than a new request.
MAX_ANSWER_CHARS = 80
# Above this many words a reply is a request in its own right, whatever it says.
MAX_ANSWER_WORDS = 8

# A question that closes the task ("Anything else?") is not a clarifying
# question: whatever the user types next is a NEW request.
_CLOSING_QUESTIONS = (
    'hanh dong khac', 'gi khac', 'gi nua', 'can gi them', 'giup gi them',
    'ho tro gi them', 'tiep theo khong', 'ban can gi', 'con gi nua',
    'anything else', 'else i can help', 'need anything', 'can i help',
    'next step', 'what next', 'how else',
    # The assistant's OWN standard closers were missing from its own list:
    # "Bạn muốn làm gì tiếp?" and the greeting "Xin chào! Bạn có cần giúp đỡ
    # gì không?" both opened the continuation path, so the next message rode
    # the previous turn instead of being routed fresh.
    'lam gi tiep', 'muon lam gi', 'can giup do gi', 'giup do gi khong',
    'help you with', 'like to do today', 'muon lam gi hom nay',
)

# Pure social openers and closers. Whatever the assistant just asked, "hello"
# is never the answer to it — and treating it as one skips the deterministic
# `greet` reply and pays for a whole LLM turn, which on a local provider is
# seconds of "Đang phản hồi…" for a word the NLU answers instantly.
#
# Deliberately NOT nlu_engine.is_conversational(), which is broader: it also
# returns True for "ok", "vâng", "được", "1" and "no" — those ARE answers to
# a clarifying question and must keep continuing the previous turn.
_SOCIAL_TURNS = (
    'hello', 'helo', 'hi', 'hey', 'yo', 'hallo',
    'good morning', 'good afternoon', 'good evening',
    'chao', 'chao ban', 'xin chao', 'alo', 'chao buoi sang',
    'thanks', 'thank you', 'thx', 'cam on', 'cam on ban', 'cam on nhe',
    'bye', 'goodbye', 'tam biet', 'see you',
)

# Replies that are answers by their very shape.
_AFFIRMATIONS = (
    'ok', 'oke', 'okay', 'yes', 'yeah', 'yep', 'sure', 'y', 'no', 'nope', 'n',
    'co', 'khong', 'k', 'dung', 'u', 'um', 'vang', 'da', 'duoc', 'roi',
    'tiep', 'tiep tuc', 'lam di', 'dong y', 'chinh xac', 'ca hai', 'both',
    'all', 'tat ca', 'toan bo', 'ca 2',
)

_ORDINAL_RE = re.compile(r'^[0-9]{1,3}$')
_LETTER_RE = re.compile(r'^[a-z]$')
# "1", "2,3", "1 và 3", "phương án 2", "option b"
_PICK_RE = re.compile(
    r'^(?:option|phuong an|cach|muc|so)?\s*[0-9a-z]'
    r'(?:\s*[,va&+]+\s*[0-9a-z])*$')


def _fold(text):
    """Diacritic-folded, whitespace-collapsed lowercase."""
    folded = vi_text.fold_diacritics(text or '').lower()
    return re.sub(r'\s+', ' ', folded).strip()


def ends_with_question(last_bot):
    """True when the assistant's last message ENDS by asking something.

    Checked on the final non-empty line, so a question mark buried mid-message
    no longer opens the continuation path.
    """
    text = (last_bot or '').strip()
    if not text:
        return False
    for line in reversed(text.split('\n')):
        line = line.strip()
        # Trailing markdown/emphasis or a stray quote must not hide the mark.
        line = line.rstrip('*_`"\'）)]】 ')
        if not line:
            continue
        return line.endswith('?') or line.endswith('？')
    return False


def is_closing_question(last_bot):
    """True for a wrap-up question ("Anything else?") rather than a real one."""
    tail = _fold(last_bot or '')[-250:]
    return any(c in tail for c in _CLOSING_QUESTIONS)


def _has_action_verb(raw):
    """True when the text carries a command verb of its own.

    Reuses the dispatcher's tables rather than restating them, so a verb added
    for classification also protects the continuation path.
    """
    try:
        from Intelligence.agents.dispatcher import (
            _ACTION_PHRASES, _ACTION_WORDS, _ACTION_COLOR_RAW, _has_phrase)
    except Exception:
        return False
    folded = _fold(raw)
    tokens = set(folded.split())
    raw_lower = re.sub(r'\s+', ' ', (raw or '').lower()).strip()
    return bool(_has_phrase(folded, _ACTION_PHRASES)
                or (tokens & set(_ACTION_WORDS))
                or any(p in raw_lower for p in _ACTION_COLOR_RAW))


def is_social_turn(raw):
    """True for a bare greeting / thanks / goodbye.

    Exact match on the folded text: "hi" alone is small talk, "hi, export the
    G sheets" is a request that happens to open politely.
    """
    return _fold(raw).rstrip('.!,?') in _SOCIAL_TURNS


def looks_like_answer(raw):
    """True when the text reads as a reply to a question, not a new command."""
    text = (raw or '').strip()
    if not text or len(text) > MAX_ANSWER_CHARS:
        return False

    folded = _fold(text)
    stripped = folded.rstrip('.!,')

    # "1", "b", "2,3", "phương án 2"
    if (_ORDINAL_RE.match(stripped) or _LETTER_RE.match(stripped)
            or _PICK_RE.match(stripped)):
        return True
    # "ok", "vâng", "toàn bộ", "both"
    if stripped in _AFFIRMATIONS:
        return True
    # A short noun phrase ("toàn bộ project", "chỉ view hiện tại") is an
    # answer; anything carrying its own verb is a new command.
    if len(stripped.split()) <= MAX_ANSWER_WORDS and not _has_action_verb(text):
        return True
    return False


def is_continuation(last_bot, raw, prev_decision, fresh_keyword_hit=False):
    """Should this turn continue the previous one instead of routing afresh?

    Args:
        last_bot: the assistant's most recent message.
        raw: the user's new message.
        prev_decision: the decision that steered the previous turn (or None).
        fresh_keyword_hit: True when the dispatcher's keyword stage matched
            this input on its own — then it is a new command regardless.
    """
    if not prev_decision or fresh_keyword_hit:
        return False
    if not (raw or '').strip():
        return False
    if not ends_with_question(last_bot):
        return False
    if is_closing_question(last_bot):
        return False
    # Small talk answers nothing — and routing it as a continuation costs a
    # full LLM turn for a word the offline NLU replies to instantly.
    if is_social_turn(raw):
        return False
    return looks_like_answer(raw)


# ─── "Which model are you using?" ────────────────────────────────────────────
# "model" is the single most overloaded word in this product: to the user it is
# either the .rvt file or the LLM behind the chat. Asked "model bạn đang dùng
# là gì", the assistant had no deterministic answer, fell through to the LLM,
# and a local model replied "Tôi chưa mở bất kỳ model Revit nào" — the wrong
# noun AND an ungrounded claim about the open document.
#
# The assistant knows its own provider and model id for certain, so this is
# answered from that fact and never sent to the LLM.

_SECOND_PERSON = ('ban', 'cau', 'may', 'minh', 'you', 'your', "you're", 'ur')
# Words that make the question unambiguously about the AI, with no need for a
# second-person pronoun.
_LLM_WORDS = ('llm', 'gpt', 'chatgpt', 'claude', 'ollama', 'lmstudio',
              'deepseek', 'provider')
_MODEL_WORDS = ('model', 'ai', 'engine')
_ASK_WORDS = ('nao', 'gi', 'what', 'which', 'who')
# A document verb settles it the other way: "model bạn vừa mở tên gì" carries
# both a pronoun and "model", but it is asking about the .rvt file.
_DOC_CONTEXT = ('mo', 'dong', 'luu', 'save', 'saved', 'open', 'opened',
                'close', 'closed', 'link', 'central', 'rvt', 'file',
                'project', 'du an', 'tai', 'load')


def asks_which_llm(text):
    """True for "which model/LLM are you running?" — never for a Revit model.

    Requires an interrogative AND either an explicit LLM word or a
    second-person pronoun, so "model nào đang mở" (a document question) and
    "bạn kiểm tra model giúp tôi" (an imperative) both stay out.
    """
    folded = _fold(text)
    if not folded:
        return False
    tokens = set(folded.replace('?', ' ').split())
    if not (tokens & set(_ASK_WORDS)) and '?' not in folded:
        return False
    llm_named = bool(tokens & set(_LLM_WORDS))
    if not llm_named and not (tokens & set(_MODEL_WORDS)):
        return False
    if not llm_named and (tokens & set(_DOC_CONTEXT)):
        return False
    return llm_named or bool(tokens & set(_SECOND_PERSON))


# ─── Learned patterns ────────────────────────────────────────────────────────
# find_learned_match() runs on Jaccard similarity >= 0.8 and used to be
# consulted BEFORE the NLU, so a remembered phrasing could take the turn with
# no way for a confident NLU read to object. The NLU knows the real tool
# catalog; a learned pattern only knows what happened to work once.

def learned_pattern_wins(learned, nlu_result, raw=None):
    """Should the learned pattern decide this turn, given what the NLU says?

    The learned pattern is used when the NLU has no opinion, or when both name
    the same intent. A disagreement goes to the NLU: it is derived from the
    live tool catalog, whereas a learned pattern can be a stale mapping to a
    tool that has since been renamed or removed.

    `raw` is the user's message. When supplied, a route the user explicitly
    marked wrong with 👎 on this phrasing is never allowed to win again —
    otherwise a bad mapping, once learned, replays forever and the down-vote
    means nothing.
    """
    if not learned or not learned.get('intent'):
        return False
    if raw:
        try:
            from Intelligence import feedback
            if feedback.is_suppressed(raw, learned.get('intent')):
                return False
        except Exception:
            pass
    if not nlu_result:
        return True
    nlu_intent = nlu_result.get('intent')
    if nlu_intent in (None, 'unknown', 'chat', 'help'):
        # No real command read — the learned pattern is the better guess,
        # unless the NLU answered authoritatively from the tool catalog.
        return not nlu_result.get('_authoritative')
    return nlu_intent == learned.get('intent')


# ─── /english-spellcheck: scan vs fix ────────────────────────────────────────

_FIX_WORDS_EN = re.compile(
    r'\b(fix|apply|correct|rename|update|replace|change)\b', re.I)
_FIX_WORDS_VI = ('sua', 'sua lai', 'sua loi', 'thay', 'doi', 'doi ten',
                 'cap nhat', 'chinh')


def wants_spellcheck_fix(skill_args):
    """True when the user asked to APPLY spelling fixes, not just scan.

    A bare scan goes to the deterministic batch pipeline, which reads every
    string; the agent playbook can only see one ai_element_filter page and
    would report full-project coverage after reading 50 notes.
    """
    args = skill_args or ''
    if _FIX_WORDS_EN.search(args):
        return True
    folded = _fold(args)
    tokens = set(folded.split())
    return any(w in tokens for w in _FIX_WORDS_VI if ' ' not in w) or \
        any(w in folded for w in _FIX_WORDS_VI if ' ' in w)

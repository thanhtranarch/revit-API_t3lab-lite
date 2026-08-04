# -*- coding: utf-8 -*-
"""
REDUCER primitive — many node results converge into one answer.

The reducer is what stops a multi-step turn from reading like a machine log.
Three sub-goals produce three answers; the user asked one question and wants
one reply, with the steps that failed named rather than silently missing.

Two modes:

    reduce()        deterministic. Headed sections in graph order, failures
                    listed explicitly, nothing invented. Always available,
                    costs nothing, and is what runs when the provider is
                    unavailable or the turn is already over budget.
    reduce_with_llm() one small call that distills the same material into
                    prose. Falls back to `reduce()` on any failure, so the
                    caller never has to handle a reducer that returned
                    nothing.

Both are bilingual: headings follow the language of the request, because a
Vietnamese turn with English section headers is exactly the kind of seam this
whole layer exists to remove.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

from Intelligence.graph.primitives import DONE, FAILED, SKIPPED, CANCELLED

# The distill call is a formatting step in front of an answer the user is
# already waiting for — it gets a short leash, like the dispatcher's classify.
_REDUCE_TIMEOUT_MS = 20000
_REDUCE_MAX_TOKENS = 900

_TEXT = {
    'vi': {
        'step':      u'Bước {}: {}',
        'failed':    u'**Không hoàn thành:** {}',
        'skipped':   u'**Đã bỏ qua** (phụ thuộc vào bước lỗi): {}',
        'cancelled': u'**Đã dừng:** {}',
        'nothing':   u'Không có bước nào hoàn thành.',
        'summary':   u'Tóm tắt',
        'reason':    u'lý do: {}',
    },
    'en': {
        'step':      u'Step {}: {}',
        'failed':    u'**Did not complete:** {}',
        'skipped':   u'**Skipped** (depended on a failed step): {}',
        'cancelled': u'**Stopped:** {}',
        'nothing':   u'No step completed.',
        'summary':   u'Summary',
        'reason':    u'reason: {}',
    },
}


def _t(lang, key):
    return _TEXT.get(lang, _TEXT['en'])[key]


class Reducer(object):
    """Converges the results of several nodes into a single reply."""

    def __init__(self, provider=None):
        self.provider = provider

    # ── deterministic ────────────────────────────────────────────────────

    def reduce(self, graph, state, node_ids=None, lang='vi'):
        """Headed concatenation of every node's result, in graph order.

        Nodes that failed, were skipped or cancelled are NAMED rather than
        dropped: a summary that quietly omits the step that did not run is the
        single most misleading thing a multi-step turn can produce.
        """
        # Dormant fallback copies are not steps of the plan — listing them
        # would report a "skipped step" for insurance that was never needed.
        ids = list(node_ids) if node_ids else [
            n.id for n in graph.nodes()
            if n.kind == 'agent' and not graph.is_fallback_target(n.id)]
        results = state.results()

        blocks = []
        done_count = 0
        step = 0
        for nid in ids:
            node = graph.get_node(nid)
            if node is None:
                continue
            step += 1
            res = results.get(nid)
            title = node.goal.strip() or nid
            if res is None:
                blocks.append(_t(lang, 'skipped').format(title))
                continue
            if res.status == DONE and (res.output or u'').strip():
                done_count += 1
                blocks.append(u'**{}**\n\n{}'.format(
                    _t(lang, 'step').format(step, title),
                    res.output.strip()))
            elif res.status == FAILED:
                detail = self._short_error(res)
                line = _t(lang, 'failed').format(title)
                if detail:
                    line += u' — ' + _t(lang, 'reason').format(detail)
                blocks.append(line)
            elif res.status == SKIPPED:
                blocks.append(_t(lang, 'skipped').format(title))
            elif res.status == CANCELLED:
                blocks.append(_t(lang, 'cancelled').format(title))
            elif (res.output or u'').strip():
                done_count += 1
                blocks.append(u'**{}**\n\n{}'.format(
                    _t(lang, 'step').format(step, title),
                    res.output.strip()))

        if not blocks:
            return _t(lang, 'nothing')
        if done_count == 0:
            # Every step failed — say so plainly instead of formatting a
            # multi-section report out of nothing but failures.
            return u'\n\n'.join([_t(lang, 'nothing')] + blocks)
        return u'\n\n---\n\n'.join(blocks)

    @staticmethod
    def _short_error(res):
        """Last meaningful line of a traceback, trimmed for a chat message."""
        err = (res.error or u'').strip()
        if not err:
            return u''
        line = [l for l in err.splitlines() if l.strip()]
        tail = line[-1].strip() if line else u''
        return tail[:160]

    # ── LLM distill ──────────────────────────────────────────────────────

    def reduce_with_llm(self, graph, state, node_ids=None, lang='vi',
                        goal=u''):
        """One small call that turns the sections into a single prose answer.

        Falls back to `reduce()` whenever the provider is missing, errors, or
        returns nothing — the deterministic version is always a valid answer,
        so a reducer failure can never cost the user their results.
        """
        base = self.reduce(graph, state, node_ids=node_ids, lang=lang)
        if self.provider is None:
            return base
        # Nothing to distill: one section is already one answer.
        sections = base.split(u'\n\n---\n\n')
        if len(sections) < 2:
            return base

        want = (u'Trả lời bằng tiếng Việt.' if lang == 'vi'
                else u'Answer in English.')
        system = (
            u"You are merging the results of several steps of ONE user "
            u"request into a single reply.\n"
            u"Rules: keep every number, element id and file path EXACTLY as "
            u"given; never add a fact that is not in the material; keep any "
            u"step that failed visible, do not smooth it away; prefer a short "
            u"lead sentence followed by the details, and keep markdown tables "
            u"as tables. {}".format(want))
        user = u"Original request: {}\n\n---\n\n{}".format(goal or graph.goal,
                                                           base)
        try:
            out = self.provider.chat([], system, user,
                                     max_tokens=_REDUCE_MAX_TOKENS,
                                     temperature=0.0,
                                     timeout_ms=_REDUCE_TIMEOUT_MS)
        except TypeError:
            # Older provider signature (no keyword budget/timeout support).
            try:
                out = self.provider.chat([], system, user, _REDUCE_MAX_TOKENS)
            except Exception:
                return base
        except Exception:
            return base
        out = (out or u'').strip()
        return out if out else base

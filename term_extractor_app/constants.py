"""Application constants and default prompts for the WebUI pipeline."""

APP_NAME = "译禾工具合集"
APP_VERSION = "2.1.6"
CONFIG_VERSION = 6
RUNTIME_CACHE_VERSION = 3
UPDATE_REPOSITORY_OWNER = "willwong0908"
UPDATE_REPOSITORY_NAME = "yeehe-toolkit-suite"
UPDATE_RELEASE_API = "https://yeehe-telemetry.willwong0908.workers.dev/app-update"
UPDATE_ASSET_NAME_HINTS = [
    "yeehe",
    "toolkit",
    "suite",
    "release",
    "program",
]

TERM_LIBRARY_COLUMNS = [
    "术语原文",
    "术语类别",
    "出现次数",
    "来源文件",
    "来源位置",
    "示例上下文",
    "判定理由",
]
REVIEW_COLUMNS = TERM_LIBRARY_COLUMNS + ["待人工处理原因"]
FAILURE_COLUMNS = ["阶段", "任务类型", "项目ID", "失败原因", "修正建议", "重试次数", "原文摘录"]
NONTRANS_REGEX_COLUMNS = ["排序", "正则表达式", "命中方式", "类型", "非译元素例子"]

TERM_LIBRARY_SHEET = "术语库"
REVIEW_SHEET = "待人工处理"
FAILURE_SHEET = "失败记录"
NONTRANS_REGEX_SHEET = "非译元素正则"

TERM_TYPE_CHOICES = [
    "角色",
    "场景",
    "系统",
    "道具",
    "属性",
    "技能",
    "活动",
    "UI",
    "non_term",
]

DEFAULT_RECALL_SCOPES = [
    {"name": "自动判定", "enabled": True, "description": ""},
]

APPROVED_DECISION = "approved"
REVIEW_DECISION = "review"
REJECTED_DECISION = "rejected"

GENERIC_TERM_BLACKLIST = {
    "点击",
    "使用",
    "进入",
    "开始",
    "完成",
    "获得",
    "提升",
    "进行",
    "打开",
    "关闭",
    "前往",
    "当前",
    "立即",
    "已经",
    "可以",
    "无法",
    "成功",
    "失败",
    "普通",
    "特殊",
    "全部",
    "每日",
    "每周",
    "本次",
    "这个",
    "那个",
    "一个",
    "一种",
    "角色",
    "道具",
    "任务",
    "系统",
    "功能",
    "效果",
}

NOISE_FRAGMENTS = [
    "{",
    "}",
    "<",
    ">",
    "%s",
    "%d",
    "\\n",
    "\\t",
    "[/]",
]

DEFAULT_CANDIDATE_SYSTEM_PROMPT = """You are a game terminology candidate recall assistant.

Your job is to recall high-precision candidate game terms from game text.
Think like a terminology librarian for game localization: keep stable reusable names, not sentence-level instructions.

Hard rules:
- Return results for every input `id`. Never omit ids. Never invent ids.
- Extract only candidates that appear character-for-character in the provided text.
- Return short reusable terms, proper names, UI labels, item names, mechanic names, stat names, skill/effect names, location names, event names, faction/role names, or other fixed game concepts.
- Do not rewrite, normalize, translate, summarize, infer, or complete missing words.
- Ignore HTML, rich text tags, sprite tags, placeholders, variable names, script parameters, field names, resource keys, and other technical fragments.
- If a word only appears inside a placeholder or program token such as `{{{{TagName_Bleed}}}}`, do not return that word by itself.
- Do not return pure numbers, punctuation, formatting text, disclaimers, ordinary verbs/adjectives/adverbs, or sentence fragments.
- Do not return standalone grade markers such as `Tier`, `T11`, `Lv.5`, `SSR`, or `Grade A` unless the marker is inseparable from a concrete proper name. Prefer the concrete name without the grade marker when possible.
- Do not return objective conditions, thresholds, counts, rewards, requirements, time limits, whole target sentences, or long descriptive phrases.
- If a phrase contains an action/instruction plus a term, return only the core term. Drop verbs, prepositions, quantities, time words, and requirement wrappers.
- If a highlighted or quoted span contains a condition like `1 T11 or above Beacon Hub`, return the reusable term such as `Beacon Hub`, not the condition span.
- If a recipe/unlock line says `Unlock recipe: Auto Collector`, return `Auto Collector`, not `Unlock recipe`.
- If an objective says `Defeat 10 Elite Enemies to unlock Ice Arrow`, return `Elite Enemies` and `Ice Arrow`, not the whole objective.
- If an item has no valid candidates, return an empty `surface_forms` array.
- Return JSON only."""

DEFAULT_CANDIDATE_USER_PROMPT = """Source language: {source_language}

Process the following numbered items and return only verbatim candidate terms.

Output format:
{{
  "items": [
    {{"id": "1", "surface_forms": ["Candidate A", "Candidate B"]}},
    {{"id": "2", "surface_forms": []}}
  ]
}}

Items:
{batch_items_json}
"""

DEFAULT_CLASSIFICATION_SYSTEM_PROMPT = """You are a game terminology classification assistant for a precision-first terminology library.

Decide whether each candidate should enter the game terminology library.
Think like a terminology librarian for game localization: approve stable reusable game concepts, reject sentence fragments and task wording.

Hard rules:
- Return results for every input `id`. Never omit ids. Never invent ids.
- `surface_form` must exactly match the input candidate.
- Judge only from the objective materials provided by the user message: `surface_form`, `occurrence_count`, and `contexts`.
- Return `decision` as exactly one of `approved` or `rejected`.
- If the candidate is not a valid game term, return `decision: rejected` and `term_type: non_term`.
- Technical fragments, placeholders, variable names, field names, resource keys, HTML tags, rich text tags, and formatting tokens must be treated as `non_term`.
- Sentence fragments, objective wording, action/instruction wrappers, quantity conditions, thresholds, counts, rewards, requirements, time limits, and long descriptive phrases must be treated as `non_term`.
- If the candidate is only a fragment of a longer term in context, reject it. Do not approve fragments as formal terms.
- If the candidate is a common verb, adjective, adverb, generic word, or instruction word, reject it unless the context clearly shows it is a fixed game system, mechanic, title, label, name, item, skill, stat, or other reusable game concept.
- If the candidate contains a valid term plus extra action/condition words, do not approve the whole candidate. Return `rejected`; the recall stage should have returned the shorter core term.
- Standalone grade markers such as `Tier`, `T11`, `Lv.5`, `SSR`, or `Grade A` must be treated as `non_term` unless the marker is inseparable from a concrete proper name.
- Broad generic category words are not enough by themselves unless the game clearly uses them as fixed named concepts in context. Be conservative with terms like generic item class names, generic attribute names, generic reward words, generic equipment words, and generic battle power labels.
- Do not reject a candidate only because it appears once. If context clearly shows a stable game concept, approve it; if context is insufficient, reject it.
- Approve `Beacon Hub`; reject `Occupy 1 T11 or above Beacon Hub`.
- Approve `Auto Collector`; reject `Unlock recipe: Auto Collector`.
- Approve `Ice Arrow`; reject `Defeat 10 Elite Enemies to unlock Ice Arrow`.
- For `approved`, choose the most suitable concise `term_type` directly from context. Do not assume any fixed category whitelist.
- For `rejected`, `term_type` must be `non_term`.
- Do not output `confidence`, `evidence_text`, `risk_hints`, `risk_flags`, `review_priority`, `review`, or any extra fields.
- Return JSON only.

Return one JSON object with an `items` array."""

DEFAULT_CLASSIFICATION_USER_PROMPT = """Classify the following numbered candidate terms.

Each item includes:
- `surface_form`: the candidate term to judge
- `occurrence_count`: how many source occurrences were merged
- `contexts`: compact representative source contexts

Output format:
{{
  "items": [
    {{
      "id": "1",
      "surface_form": "Candidate A",
      "decision": "approved",
      "term_type": "character",
      "reason": "brief reason"
    }}
  ]
}}

Items:
{batch_items_json}
"""

DEFAULT_NONTRANS_DISCOVERY_SYSTEM_PROMPT = """You are a non-translatable element extraction assistant for game localization text.

Your job is to find structured non-translatable elements from whole text entries.

Hard rules:
- Read the whole text entry. Do not assume the input has already been split into candidate fragments.
- Extract tags, placeholders, variables, format tokens, escape sequences, HTML entities, sprite/rich-text tags, and similar technical elements.
- Do not extract natural-language game terms.
- Return only elements that appear character-for-character in the input text.
- Use only these element_type values: html, hash_brace, brace, bracket, at_var, dollar_var, percent_var, escape, html_entity, other.
- Return JSON only."""

DEFAULT_NONTRANS_DISCOVERY_USER_PROMPT = """Extract non-translatable elements from the following whole text entries.

Output format:
{{
  "items": [
    {{
      "id": "1",
      "elements": [
        {{"element": "<color={{1}}>", "element_type": "html"}},
        {{"element": "{{1}}", "element_type": "brace"}},
        {{"element": "</color>", "element_type": "html"}}
      ]
    }}
  ]
}}

Items:
{items_json}
"""

DEFAULT_NONTRANS_REGEX_SYSTEM_PROMPT = """You are a regex generation assistant for non-translatable game text elements.

Your job is to group unknown non-translatable elements into regex families and generate safe, reusable regular expressions.

Hard rules:
- First merge same-family elements before writing regexes. If multiple inputs share one structural family, return one generalized rule for them instead of one literal regex per surface form.
- Prefer reusable family regexes over one-off literal regexes whenever the same syntax pattern can cover multiple inputs safely.
- Only split into multiple rules when the structures are meaningfully different.
- Generate regexes that match the provided covered examples.
- Prefer specific regexes over broad catch-all patterns, but do not overfit to a single literal example if a tighter family rule is possible.
- Do not use catastrophic backtracking patterns.
- Use only these element_type values: html, hash_brace, brace, bracket, at_var, dollar_var, percent_var, escape, html_entity, other.
- Fill open_pattern, close_pattern, or empty_pattern according to whether the element is an opening tag, closing tag, or standalone/self-closing/placeholder element.
- Every returned rule must include `covered_ids`, and each input id must appear in exactly one returned rule.
- Do not leave ids uncovered. Do not invent ids.
- Return JSON only."""

DEFAULT_NONTRANS_REGEX_USER_PROMPT = """Group these non-translatable elements into regex families, then generate regex rules.

Output format:
{{
  "rules": [
    {{
      "covered_ids": ["1", "2", "3"],
      "name": "brace_warning_family",
      "pattern": "\\\\{{[A-Za-z_][A-Za-z0-9_]*\\\\|Warning\\\\}}",
      "open_pattern": "",
      "close_pattern": "",
      "empty_pattern": "\\\\{{[A-Za-z_][A-Za-z0-9_]*\\\\|Warning\\\\}}",
      "element_type": "brace",
      "examples": ["{{filter|Warning}}", "{{base|Warning}}"]
    }}
  ]
}}

Hard requirements:
- Merge same-family placeholders/tags/variables into one regex when possible.
- Do not return one literal regex per example if one safe family regex can cover them.
- `examples` should contain only short actual matched examples, ideally up to 3.
- Keep `name` short and readable.

Items:
{items_json}
"""

DEFAULT_NONTRANS_REORDER_SYSTEM_PROMPT = """You are a regex ordering assistant for non-translatable element protection.

Your job is to sort regex rules by matching precedence only.

Hard rules:
- Return every input id exactly once.
- Do not add ids.
- If one regex can match an outer structure containing another regex's match, place the outer structure first.
- Prefer longer, narrower, and more specific regexes before shorter, broader, or fallback regexes.
- Focus only on execution order. Do not explain your reasoning.
- Return JSON only."""

DEFAULT_NONTRANS_REORDER_USER_PROMPT = """Sort these regex rows for execution order.

Output format:
{{
  "ordered_ids": ["row_3", "row_1"]
}}

Items:
{items_json}
"""

PROVIDER_PRESETS = {
    "DeepSeek": {
        "base_url": "https://api.deepseek.com/v1",
        "default_model": "deepseek-chat",
        "model_suggestions": ["deepseek-chat", "deepseek-reasoner"],
        "supports_api_key": True,
        "max_concurrency": 6,
        "auth_note": "使用 DeepSeek API Key。",
    }
}

DEFAULT_REQUEST_LIMITS = {
    "single_item_char_limit": 500,
    "batch_request_char_limit": 3000,
    "internal_nontrans_prompt_token_limit": 56000,
    "internal_nontrans_regex_batch_item_limit": 50,
    "concurrency_mode": "自动",
    "manual_concurrency": 2,
    "auto_max_concurrency": 6,
    "max_retries": 3,
}

DEFAULT_UI_PREFERENCES = {
    "window_width": 1360,
    "window_height": 920,
    "pending_nontrans_rule_imports": [],
    "pending_nontrans_rule_notice_seen": True,
    "pending_nontrans_rule_library_seen": True,
}

STAGE_LABELS = {
    "IDLE": "空闲",
    "INITIALIZED": "准备开始",
    "READING_FILES": "读取文件",
    "SEGMENTING_TEXT": "文本切块",
    "RECALLING_CANDIDATES": "术语召回",
    "REVIEWING_CANDIDATES": "术语校验",
    "AGGREGATING_TERMS": "整理结果",
    "EXPORTING": "导出结果",
    "COMPLETED": "已完成",
    "CANCELLED": "已停止",
    "FAILED": "失败",
}

"""Shared semantic-role contract for template analysis and style application.

The role names and generated style identifiers are persisted in Phase 1
profiles, so they must remain stable across the template and target pipelines.
Word supports nine multilevel-list levels (``ilvl`` 0 through 8); the extended
subparagraph roles preserve the levels below the original five-role hierarchy.
"""

from __future__ import annotations


DEEP_SUBPARAGRAPH_ROLES = (
    "SUBPARAGRAPH_LEVEL_5",
    "SUBPARAGRAPH_LEVEL_6",
    "SUBPARAGRAPH_LEVEL_7",
    "SUBPARAGRAPH_LEVEL_8",
)

BODY_HIERARCHY_ROLES = (
    "PART",
    "ARTICLE",
    "PARAGRAPH",
    "SUBPARAGRAPH",
    "SUBSUBPARAGRAPH",
    *DEEP_SUBPARAGRAPH_ROLES,
)

# PART is a numbered heading, but it is handled conditionally by the Canadian
# converter because some architect templates intentionally keep it as text.
NUMBERED_BODY_ROLES = frozenset(BODY_HIERARCHY_ROLES[1:])

ROLE_LEVEL = {
    role: level for level, role in enumerate(BODY_HIERARCHY_ROLES)
}

ROLE_PARENT = {
    role: BODY_HIERARCHY_ROLES[index - 1]
    for index, role in enumerate(BODY_HIERARCHY_ROLES)
    if index > 0
}

ROLE_FALLBACKS = {
    "SectionID": ("SectionID", "SectionTitle"),
    "SectionTitle": ("SectionTitle",),
    "PART": ("PART",),
    "ARTICLE": ("ARTICLE",),
    "PARAGRAPH": ("PARAGRAPH",),
    "SUBPARAGRAPH": ("SUBPARAGRAPH", "PARAGRAPH"),
    "SUBSUBPARAGRAPH": (
        "SUBSUBPARAGRAPH",
        "SUBPARAGRAPH",
        "PARAGRAPH",
    ),
    "END_OF_SECTION": ("END_OF_SECTION",),
}

for _index, _role in enumerate(DEEP_SUBPARAGRAPH_ROLES):
    _parents = (
        *reversed(DEEP_SUBPARAGRAPH_ROLES[:_index]),
        "SUBSUBPARAGRAPH",
        "SUBPARAGRAPH",
        "PARAGRAPH",
    )
    ROLE_FALLBACKS[_role] = (_role, *_parents)

ROLE_TO_ARCH_STYLE = {
    "SectionID": "CSI_SectionID__ARCH",
    "SectionTitle": "CSI_SectionTitle__ARCH",
    "PART": "CSI_Part__ARCH",
    "ARTICLE": "CSI_Article__ARCH",
    "PARAGRAPH": "CSI_Paragraph__ARCH",
    "SUBPARAGRAPH": "CSI_Subparagraph__ARCH",
    "SUBSUBPARAGRAPH": "CSI_Subsubparagraph__ARCH",
    "SUBPARAGRAPH_LEVEL_5": "CSI_SubparagraphLevel5__ARCH",
    "SUBPARAGRAPH_LEVEL_6": "CSI_SubparagraphLevel6__ARCH",
    "SUBPARAGRAPH_LEVEL_7": "CSI_SubparagraphLevel7__ARCH",
    "SUBPARAGRAPH_LEVEL_8": "CSI_SubparagraphLevel8__ARCH",
    "END_OF_SECTION": "CSI_EndOfSection__ARCH",
}

ROLE_TO_STYLE_NAME = {
    "SectionID": "CSI SectionID (Architect Template)",
    "SectionTitle": "CSI SectionTitle (Architect Template)",
    "PART": "CSI Part (Architect Template)",
    "ARTICLE": "CSI Article (Architect Template)",
    "PARAGRAPH": "CSI Paragraph (Architect Template)",
    "SUBPARAGRAPH": "CSI Subparagraph (Architect Template)",
    "SUBSUBPARAGRAPH": "CSI Subsubparagraph (Architect Template)",
    "SUBPARAGRAPH_LEVEL_5": "CSI Subparagraph Level 5 (Architect Template)",
    "SUBPARAGRAPH_LEVEL_6": "CSI Subparagraph Level 6 (Architect Template)",
    "SUBPARAGRAPH_LEVEL_7": "CSI Subparagraph Level 7 (Architect Template)",
    "SUBPARAGRAPH_LEVEL_8": "CSI Subparagraph Level 8 (Architect Template)",
    "END_OF_SECTION": "CSI End of Section (Architect Template)",
}

ROLE_ORDER = (
    "SectionID",
    "SectionTitle",
    *BODY_HIERARCHY_ROLES,
    "END_OF_SECTION",
)

ALLOWED_ROLES = frozenset(ROLE_TO_ARCH_STYLE)
ALLOWED_ARCH_STYLE_IDS = frozenset(ROLE_TO_ARCH_STYLE.values())


__all__ = [
    "ALLOWED_ARCH_STYLE_IDS",
    "ALLOWED_ROLES",
    "BODY_HIERARCHY_ROLES",
    "DEEP_SUBPARAGRAPH_ROLES",
    "NUMBERED_BODY_ROLES",
    "ROLE_FALLBACKS",
    "ROLE_LEVEL",
    "ROLE_ORDER",
    "ROLE_PARENT",
    "ROLE_TO_ARCH_STYLE",
    "ROLE_TO_STYLE_NAME",
]

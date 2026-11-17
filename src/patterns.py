"""Регулярные выражения RU/EN и связанные хелперы."""

import re

from . import i18n


NUM_PART = r'\d+(?:\.\d+)?'
DASH = r'[-\u2010\u2011\u2012\u2013\u2014\u2015\u2212]'

NUM_LIST = (NUM_PART +
            r'(?:\s*' + DASH + r'\s*' + NUM_PART + r')?' +
            r'(?:\s*(?:,|and|и)\s*' + NUM_PART + r')*')

CAPTION_RE_RU = re.compile(
    r'^\s*(Рисунок|Рис\.?|Таблица|Табл\.?)\s*(?:№\s*)?(' + NUM_PART + r'[а-я]?)\s*'
    r'(?:[–—\-−.:]\s*|$)',
    re.IGNORECASE,
)
CAPTION_RE_EN = re.compile(
    r'^\s*(Figure|Fig\.?|Table|Tab\.?|TABLE)\s*(' + NUM_PART + r'[a-z]?)\s*'
    r'(?:[–—\-−.:]\s*|$)',
    re.IGNORECASE,
)

CYR = r'[А-Яеа-яе]'
FIG_WORDS_RU = r'(?:рисун' + CYR + r'*|рис\.?)'
TAB_WORDS_RU = r'(?:таблиц' + CYR + r'*|табл\.?)'
EQ_WORDS_RU  = r'(?:формул' + CYR + r'*|выражени' + CYR + r'*|соотношени' + CYR + r'*|равенств' + CYR + r'*)'

FIG_WORDS_EN = r'(?:Figures?|Figs?\.?)'
TAB_WORDS_EN = r'(?:Tables?|Tab\.?)'
EQ_WORDS_EN  = r'(?:Equations?|Eqs?\.?|Expressions?|Relations?|Equalities?)'


def make_ref_regex(noun_pattern):
    return re.compile(
        r'(?<![\wА-Яеа-яе])(?P<noun>' + noun_pattern + r')\s*(?::\s*)?(?:\(\s*)?'
        r'(?P<num>' + NUM_LIST + r')(?:\s*\))?(?![А-Яеа-яе\d\-])',
        re.IGNORECASE)


FIG_REF_RE_RU = make_ref_regex(FIG_WORDS_RU)
TAB_REF_RE_RU = make_ref_regex(TAB_WORDS_RU)
EQ_REF_RE_RU  = make_ref_regex(EQ_WORDS_RU)

FIG_REF_RE_EN = make_ref_regex(FIG_WORDS_EN)
TAB_REF_RE_EN = make_ref_regex(TAB_WORDS_EN)
EQ_REF_RE_EN  = make_ref_regex(EQ_WORDS_EN)

SRC_REF_RE = re.compile(r'\[\s*([^\[\]]+?)\s*\]')

BIB_HEADING_RE_RU = re.compile(
    r'^\s*(?:\d+[\.\s]*)?(?:'
    r'библиографическ' + CYR + r'*\s+список(?:\s+литературы|\s+источников)?|'
    r'список\s+(?:использованн|используем)' + CYR + r'*\s+(?:литературы|источников)(?:\s+и\s+литературы)?|'
    r'список\s+литературы|'
    r'список\s+источников|'
    r'использованн' + CYR + r'*\s+(?:литература|источники)|'
    r'литература|'
    r'библиография|'
    r'источники'
    r')\s*[:\.]?\s*$',
    re.IGNORECASE)

BIB_HEADING_RE_EN = re.compile(
    r'^\s*(?:\d+[\.\s]*)?(?:'
    r'References|Bibliography|Literature Cited|Works Cited|'
    r'Reference List|Sources'
    r')\s*[:\.]?\s*$',
    re.IGNORECASE)

BIB_END_RE = re.compile(
    r'^\s*(приложени' + CYR + r'*|заключение|введение|оглавление|содержание|'
    r'Appendix|Conclusion|Introduction|Contents|Acknowledgments)\s*$',
    re.IGNORECASE)

TOC_LINE_RE = re.compile(r'(?:\.{3,}|\t|\s{4,}|№\s*стр\.)\s*\d*\s*$', re.IGNORECASE)
MATH_SYMBOLS = '=×÷∫∑√±∆∇≈≠≤≥·∓−'


def get_caption_re():
    return CAPTION_RE_EN if i18n.current_doc_lang == 'en' else CAPTION_RE_RU


def get_fig_ref_re():
    return FIG_REF_RE_EN if i18n.current_doc_lang == 'en' else FIG_REF_RE_RU


def get_tab_ref_re():
    return TAB_REF_RE_EN if i18n.current_doc_lang == 'en' else TAB_REF_RE_RU


def get_eq_ref_re():
    return EQ_REF_RE_EN if i18n.current_doc_lang == 'en' else EQ_REF_RE_RU


def get_bib_heading_re():
    return BIB_HEADING_RE_EN if i18n.current_doc_lang == 'en' else BIB_HEADING_RE_RU


def normalize_num(num_str):
    s = str(num_str).replace('\xa0', ' ').replace('\u202f', ' ')
    s = re.sub(r'[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]', '-', s)
    s = re.sub(r'\s+', '', s)
    if s.endswith('.'):
        s = s[:-1]
    return s
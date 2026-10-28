import re

# Earliest-mentioned clinical keyword wins — this matches how OSCAR reason text is written
# ("NP: <reason1>/<reason2>/...") regardless of odd punctuation (commas, parens, dashes).
REFRACTIVE_KW = r'\bcat\b|\bcataract\b|\brle\b'
OTHER_KW = (r'\bglc\b|\bglaucoma\b|\bretina\b|\bcornea\b|\bgeneral\b|\bdry\s*eye\b|\blid\b|'
            r'\buveitis\b|\bpterygium\b|\bpco\b|\boculoplastics?\b|\bneuro\b|\bvitreous\b|'
            r'\bamd\b|\barmd\b|\bdiabetic\b|\bdr\b|\bfloaters?\b|\bforeign\s*body\b|'
            r'\blpi\b|\bslt\b|\biritis\b|\bplastics\b|\bscleritis\b|\bprk\b|\bick\b|'
            r'\bicl\b|\brcs\b|\blal\b')

def _earliest_match(text, pattern):
    m = re.search(pattern, text, re.IGNORECASE)
    return m.start() if m else None

def is_refractive(reason):
    """True if the EARLIEST-mentioned clinical reason in the text is cataract/RLE-related,
    i.e. it's the primary reason, not a secondary item later in a list."""
    s = str(reason)
    r_pos = _earliest_match(s, REFRACTIVE_KW)
    if r_pos is None:
        return False
    o_pos = _earliest_match(s, OTHER_KW)
    if o_pos is None:
        return True
    return r_pos < o_pos

def primary_hit(reason):
    """Debug helper: which keyword matched earliest, and where."""
    s = str(reason)
    r_pos = _earliest_match(s, REFRACTIVE_KW)
    o_pos = _earliest_match(s, OTHER_KW)
    if r_pos is None and o_pos is None:
        return ('none', None)
    if o_pos is None or (r_pos is not None and r_pos < o_pos):
        return ('refractive', r_pos)
    return ('other', o_pos)

if __name__ == '__main__':
    import pandas as pd
    df = pd.read_pickle('all3.pkl')
    jul = df[df.date >= pd.Timestamp('2025-07-01')].copy()
    consults = jul[jul.type.isin(['Consult MD', 'Consult OD'])].copy()
    consults['refractive'] = consults['reason'].apply(is_refractive)
    print('Total consults Jul2025+:', len(consults))
    print(consults.refractive.value_counts())

    # re-check the 9 RLE-mismatch rows from before
    check = [
        '1/12 K check biometry then dilate, RLE consult + will need OPD pentacam, IOLmast',
        'SELF-REFERRAL - C1 for MRx PRE-CONSULT (RLE vs. PRK) ND pt has to leave by 12:00',
        'SELF REFERRAL: PRK/ICK/RLE ? Sadaf to tech, do MRx - complete VF, OCT ONH, OCT M',
        'NP: Cataract OD/RLE OS/cornea dil  YAV c1 tg',
        'c1 tgNP: LPI vs RLE',
        'c2 tg NP: LPI vs RLE dilate',
        'NP: (glc s/p RLE) ND c6 dp',
        'ND PREV PT: Glc no dil  *pt interested in RLE*  c3gt',
        'NP: LPI vs RLE no dil //',
    ]
    print()
    for r in check:
        print(is_refractive(r), primary_hit(r), '|', r)

    # no-keyword-at-all rows (fall back to 'other' / unclassified) — how many?
    no_kw = consults[consults['reason'].apply(lambda r: primary_hit(r)[0] == 'none')]
    print()
    print('No recognizable keyword at all:', len(no_kw))
    print(no_kw['reason'].head(20).tolist())

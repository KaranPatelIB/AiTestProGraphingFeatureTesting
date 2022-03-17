"""Spelling Corrector in Python 3; see http://norvig.com/spell-correct.html

Copyright (c) 2007-2016 Peter Norvig
MIT license: www.opensource.org/licenses/mit-license.php
"""

# Spelling Corrector

import re
from collections import Counter

WORDS = Counter(re.findall(r'\w+', open('big.txt').read().lower()))


class norvigSpellChecker:

    @staticmethod
    def P(word, N=sum(WORDS.values())):
        # "Probability of `word`."
        return WORDS[word] / N

    @staticmethod
    def correction(word):
        # "Most probable spelling correction for word."
        return max(norvigSpellChecker.candidates(word), key=norvigSpellChecker.P)

    @staticmethod
    def candidates(word):
        # "Generate possible spelling corrections for word."
        return norvigSpellChecker.known([word]) or norvigSpellChecker.known(
            norvigSpellChecker.edits1(word)) or norvigSpellChecker.known(norvigSpellChecker.edits2(word)) or [word]

    @staticmethod
    def known(words):
        # "The subset of `words` that appear in the dictionary of WORDS."
        return set(w for w in words if w in WORDS)

    @staticmethod
    def edits1(word):
        # "All edits that are one edit away from `word`."
        letters = 'abcdefghijklmnopqrstuvwxyz'
        splits = [(word[:i], word[i:]) for i in range(len(word) + 1)]
        deletes = [L + R[1:] for L, R in splits if R]
        transposes = [L + R[1] + R[0] + R[2:] for L, R in splits if len(R) > 1]
        replaces = [L + c + R[1:] for L, R in splits if R for c in letters]
        inserts = [L + c + R for L, R in splits for c in letters]
        return set(deletes + transposes + replaces + inserts)

    @staticmethod
    def edits2(word):
        # "All edits that are two edits away from `word`."
        return (e2 for e1 in norvigSpellChecker.edits1(word) for e2 in norvigSpellChecker.edits1(e1))


print(norvigSpellChecker.correction('teesttt'))

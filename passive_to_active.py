import spacy

nlp = spacy.load("en_core_web_sm")

# ------------------ Irregular verbs ------------------
IRREGULAR_PAST = {
    "be": "was",
    "begin": "began",
    "break": "broke",
    "bring": "brought",
    "build": "built",
    "buy": "bought",
    "catch": "caught",
    "choose": "chose",
    "come": "came",
    "do": "did",
    "drink": "drank",
    "drive": "drove",
    "eat": "ate",
    "fall": "fell",
    "feel": "felt",
    "find": "found",
    "give": "gave",
    "go": "went",
    "have": "had",
    "hear": "heard",
    "keep": "kept",
    "know": "knew",
    "leave": "left",
    "make": "made",
    "meet": "met",
    "pay": "paid",
    "put": "put",
    "read": "read",
    "run": "ran",
    "say": "said",
    "send": "sent",
    "sit": "sat",
    "speak": "spoke",
    "stand": "stood",
    "take": "took",
    "teach": "taught",
    "tell": "told",
    "think": "thought",
    "write": "wrote",
    "sing": "sang",
    "swim": "swam",
    "ring": "rang",
    "see": "saw",
    "get": "got",
    "forget": "forgot",
    "cut": "cut",
    "hit": "hit",
    "set": "set",
    "shut": "shut",
    "spread": "spread",
    "beat": "beat",
    "become": "became",
    "bite": "bit",
    "blow": "blew",
    "draw": "drew",
    "grow": "grew",
    "hide": "hid",
    "throw": "threw",
    "wear": "wore",
    "tear": "tore",
    "weep": "wept",
    "win": "won",
    "feed": "fed",
    "lead": "led",
    "light": "lit",
    "lose": "lost",
    "sell": "sold",
    "shoot": "shot",
    "slide": "slid",
    "stick": "stuck",
    "sting": "stung",
    "strike": "struck",
    "sweep": "swept",
    "swing": "swung",
    "wake": "woke",
    "bend": "bent",
    "bleed": "bled",
    "breed": "bred",
    "deal": "dealt",
    "dig": "dug",
    "hang": "hung",
    "kneel": "knelt",
    "lean": "leant",
    "leap": "leapt",
    "mean": "meant",
    "shine": "shone",
    "smell": "smelt",
    "spell": "spelt",
    "spill": "spilt",
    "spoil": "spoilt",
    "burn": "burnt",
}

PRONOUN_OBJ = {
    "i": "me",
    "we": "us",
    "you": "you",
    "he": "him",
    "she": "her",
    "it": "it",
    "they": "them",
}

MODALS = ["can", "could", "may", "might", "shall", "should", "will", "would", "must"]


# ------------------ Helper functions ------------------
def get_phrase(token):
    return " ".join([t.text for t in sorted(token.subtree, key=lambda x: x.i)])


def to_past(verb):
    verb_lower = verb.lower()
    if verb_lower in IRREGULAR_PAST:
        return IRREGULAR_PAST[verb_lower]

    if verb.endswith("e"):
        return verb + "d"
    elif verb.endswith("y") and len(verb) > 1 and verb[-2] not in "aeiou":
        return verb[:-1] + "ied"
    elif (
        len(verb) > 2
        and verb[-1] not in "aeiou"
        and verb[-2] in "aeiou"
        and verb[-3] not in "aeiou"
    ):
        return verb + verb[-1] + "ed"
    else:
        return verb + "ed"


def vbn_to_vbg(verb):
    base_form = verb
    for base, past in IRREGULAR_PAST.items():
        if past == verb.lower():
            base_form = base
            break

    if base_form.endswith("e") and base_form not in ["be", "see", "flee", "agree"]:
        return base_form[:-1] + "ing"
    elif base_form.endswith("ie"):
        return base_form[:-2] + "ying"
    elif base_form.endswith("ed"):
        return base_form[:-2] + "ing"
    elif (
        len(base_form) > 2
        and base_form[-1] not in "aeiou"
        and base_form[-2] in "aeiou"
        and base_form[-3] not in "aeiou"
    ):
        return base_form + base_form[-1] + "ing"
    else:
        return base_form + "ing"


def normalize_object(text):
    words = text.split()
    if (
        len(words) > 1
        and words[0][0].isupper()
        and words[0].lower() not in ["alice", "bob", "john", "mary", "ceo"]
    ):
        words[0] = words[0].lower()
    return " ".join(words)


def get_subject_verb_agreement(subject, base_verb):
    subject_lower = subject.lower().strip()
    if subject_lower == "i":
        if base_verb == "be":
            return "am"
        elif base_verb in ["have", "has"]:
            return "have"
        else:
            return base_verb

    if subject_lower in ["he", "she", "it", "someone", "everyone", "anyone", "nobody"]:
        if base_verb == "be":
            return "is"
        elif base_verb == "have":
            return "has"
        elif base_verb.endswith("y") and base_verb[-2] not in "aeiou":
            return base_verb[:-1] + "ies"
        elif base_verb.endswith(("s", "x", "z", "ch", "sh")):
            return base_verb + "es"
        else:
            return base_verb + "s"

    if subject_lower in ["we", "you", "they"] or "," in subject or " and " in subject:
        if base_verb == "be":
            return "are"
        elif base_verb == "have":
            return "have"
        else:
            return base_verb

    return base_verb


# ------------------ Core passive-to-active processing ------------------
def process_passive_clause(clause_doc):
    agent = None
    patient = None
    main_verb = None
    aux_tokens = []
    negations = []
    time_phrases = []

    for token in clause_doc:
        if token.dep_ == "agent":
            agent = " ".join(
                [
                    t.text
                    for t in sorted(token.subtree, key=lambda x: x.i)
                    if t.text.lower() != "by"
                ]
            )
        elif token.dep_ == "nsubjpass":
            patient = get_phrase(token)
        elif token.dep_ in ("aux", "auxpass") or token.tag_ == "MD":
            aux_tokens.append(token.text.lower())
        elif token.pos_ == "VERB" and token.dep_ == "ROOT":
            main_verb = token.lemma_
        elif token.dep_ == "neg":
            negations.append(token.text)
        elif token.dep_ in ["advmod", "npadvmod"] and token.head.pos_ == "VERB":
            time_phrases.append(token.text)

    if not agent:
        agent = "someone"
    if not patient:
        for token in clause_doc:
            if token.dep_ == "nsubj":
                patient = get_phrase(token)
                break
        if not patient:
            patient = "something"
    if not main_verb:
        return clause_doc.text

    # Pronoun handling
    first_word = patient.split()[0].lower()
    if first_word in PRONOUN_OBJ:
        patient_words = patient.split()
        patient_words[0] = PRONOUN_OBJ[first_word]
        patient = " ".join(patient_words)

    patient = normalize_object(patient)

    negation = " ".join(negations) + " " if negations else ""
    verb_phrase = ""

    # Modal verbs
    modals = [a for a in aux_tokens if a in MODALS]
    if modals:
        modal = modals[0]
        verb_phrase = f"{negation}{modal} {main_verb}"
    elif "been" in aux_tokens:
        have_verbs = [a for a in aux_tokens if a in ["has", "have", "had"]]
        if have_verbs:
            have_verb = have_verbs[0]
            verb_phrase = f"{negation}{have_verb} {main_verb}"
        else:
            verb_phrase = f"{negation}{main_verb}"
    elif "being" in aux_tokens:
        be_verb = get_subject_verb_agreement(agent, "be")
        present_participle = vbn_to_vbg(main_verb)
        verb_phrase = f"{negation}{be_verb} {present_participle}"
    else:
        be_forms = [
            a for a in aux_tokens if a in ["am", "is", "are", "was", "were", "be"]
        ]
        if be_forms:
            be_verb = be_forms[0]
            if be_verb in ["am", "is", "are"]:
                active_verb = get_subject_verb_agreement(agent, main_verb)
                verb_phrase = f"{negation}{active_verb}"
            elif be_verb in ["was", "were"]:
                past_verb = to_past(main_verb)
                verb_phrase = f"{negation}{past_verb}"
            else:
                verb_phrase = f"{negation}{main_verb}"
        else:
            past_verb = to_past(main_verb)
            verb_phrase = f"{negation}{past_verb}"

    # Additional phrases
    additional_phrases = []
    included_tokens = set()
    for token in clause_doc:
        if token.i in included_tokens:
            continue
        if token.dep_ in ["agent", "nsubjpass", "aux", "auxpass", "ROOT"]:
            continue
        if token.dep_ in ["dobj", "attr", "oprd", "acomp"]:
            phrase = get_phrase(token)
            if phrase.lower() not in [agent.lower(), patient.lower()]:
                additional_phrases.append(phrase)
                for t in token.subtree:
                    included_tokens.add(t.i)
        elif token.dep_ == "prep" and token.text.lower() != "by":
            prep_phrase = get_phrase(token)
            additional_phrases.append(prep_phrase)
            for t in token.subtree:
                included_tokens.add(t.i)
        elif token.dep_ in ["advmod", "npadvmod"] and token.text not in time_phrases:
            phrase = get_phrase(token)
            additional_phrases.append(phrase)
            for t in token.subtree:
                included_tokens.add(t.i)

    # Conditional clauses
    conditional_phrase = ""
    for token in clause_doc:
        if token.text.lower() in ["if", "when", "unless"] and token.dep_ == "mark":
            conditional_clause = get_phrase(token.head)
            conditional_phrase = f" {token.text} {conditional_clause}"
            break

    # Construct active sentence
    agent_cap = agent[0].upper() + agent[1:] if len(agent) > 0 else agent
    active_sentence = f"{agent_cap} {verb_phrase} {patient} {' '.join(additional_phrases)}{conditional_phrase}".strip()
    active_sentence = " ".join(active_sentence.split())

    for keyword in ["before being", "after being", "when being", "while being"]:
        if keyword in active_sentence.lower():
            parts = active_sentence.split()
            for i in range(len(parts) - 2):
                if (
                    parts[i].lower() == keyword.split()[0]
                    and parts[i + 1].lower() == "being"
                ):
                    ing_form = vbn_to_vbg(parts[i + 2])
                    parts[i + 1] = ing_form
                    del parts[i + 2]
                    break
            active_sentence = " ".join(parts)

    return active_sentence


# ------------------ Single sentence / clause ------------------
def passive_to_active(sentence):
    doc = nlp(sentence)
    if " and " in sentence.lower():
        clauses = []
        current_clause = []

        for token in doc:
            if token.text.lower() == "and" and token.dep_ == "cc":
                if current_clause:
                    clauses.append(" ".join([t.text for t in current_clause]))
                    current_clause = []
            else:
                current_clause.append(token)

        if current_clause:
            clauses.append(" ".join([t.text for t in current_clause]))

        active_clauses = []
        for clause in clauses:
            clause_doc = nlp(clause)
            active_clause = process_passive_clause(clause_doc)
            active_clauses.append(active_clause)

        return " and ".join(active_clauses) + "."
    else:
        result = process_passive_clause(doc)
        if not result.endswith("."):
            result += "."
        return result

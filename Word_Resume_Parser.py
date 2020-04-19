import docx2txt
import re
from spacy.matcher import Matcher
import pandas as pd
import spacy
import string

# Read the content from resume (word file)
def extract_text_from_doc(doc_path):
    temp = docx2txt.process(R_path)
    text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
    return ' '.join(text)

# This Function will extract Mobile Number
def extract_mobile_number(text):
    phone = re.findall(re.compile(r'(?:(?:\+?([1-9]|[0-9][0-9]|[0-9][0-9][0-9])\s*(?:[.-]\s*)?)?(?:\(\s*([2-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9])\s*\)|([0-9][1-9]|[0-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9]))\s*(?:[.-]\s*)?)?([2-9]1[02-9]|[2-9][02-9]1|[2-9][02-9]{3})\s*(?:[.-]\s*)?([0-9]{4})(?:\s*(?:#|x\.?|ext\.?|extension)\s*(\d+))?'), text)
    if phone:
        number = ''.join(phone[0])
        if len(number) >= 10:
            return '+' + number
        else:
            return number

# This Function will extract Email ID
def extract_email(email):
    email = re.findall("([^@|\s]+@[^@]+\.[^@|\s]+)", email)
    if email:
        try:
            return email[0].split()[0].strip(';')
        except IndexError:
            return None


# load pre-trained model
nlp = spacy.load('en_core_web_sm')
# initialize matcher with a vocab
matcher = Matcher(nlp.vocab)

# This Function will extract Name of the Candidate
def extract_name(text):
    nlp_text = nlp(text)
    # First name and Last name are always Proper Nouns
    pattern = [[{'POS': 'PROPN'}], [{'POS': 'PROPN'}] ]
    matcher.add('NAME', None, *pattern)
    matches = matcher(nlp_text)
    for matches_id, start, end in matches[0:3]:
        string_id = nlp.vocab.strings[matches_id]
        span = nlp_text[start:end]
        print(span.text, end =' ')

# This Function will extract Skills of the Candidate
def extract_skills(text):
    nlp_text = nlp(text)

    # removing stop words and implementing word tokenization
    tokens = [token.text for token in nlp_text if not token.is_stop]

    # removing punctuations from the lists
    tokens = [c for c in tokens if c not in string.punctuation]

    # reading the csv file
    data = pd.read_csv("F:\skills.csv")

    # extract values
    skills = list(data.columns.values)
    # print(skills)

    skillset = []
    # check for one-grams (example: python)
    for token in tokens:
        if token.lower() in skills:
            skillset.append(token)

    # # check for bi-grams and tri-grams (example: machine learning)
    # # for token in noun_chunks:
    # #     token = token.text.lower().strip()
    # #     if token in skills:
    # #         skillset.append(token)

    return set([i.capitalize() for i in set([i.lower() for i in skillset])])


# Program Will start execute from below
if __name__ == "__main__":
    R_path = input("Enter your word file resume name with extension")
    data = extract_text_from_doc(R_path)
    print(extract_name(data))
    print(f"Email ID : {extract_email(data)}")
    print(f"Mobile No. : {extract_mobile_number(data)}")
    print(f"Skills : {extract_skills(data)}")


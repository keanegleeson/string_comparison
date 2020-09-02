import pandas as pd

'''
Option 1)
Method explained here: 
https://towardsdatascience.com/calculating-string-similarity-in-python-276e18a7d33a
uses sklearn, might require more cleaning and prep than what I need
import string
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import CountVectorizer
# from nltk.corpus import stopwords 
'''

'''
Option 2) -https://stackoverflow.com/questions/55162668/calculate-similarity-between-list-of-words

'''

#script for option 2
def word2vec(word):
    from collections import Counter
    from math import sqrt

    # count the characters in word
    cw = Counter(word)
    # precomputes a set of the different characters
    sw = set(cw)
    # precomputes the "length" of the word vector
    lw = sqrt(sum(c*c for c in cw.values()))

    # return a tuple
    return cw, sw, lw

def cosdis(v1, v2):
    # which characters are common to the two words?
    common = v1[1].intersection(v2[1])
    # by definition of cosine distance we have
    return sum(v1[0][ch]*v2[0][ch] for ch in common)/v1[2]/v2[2]

'''
Bringing in utility names from spreadsheet which I had brought into a Qlik app originally
One column has all the utility names in the egrid dataset

Other col had utility names for Digital Realty held in RA

The idea behind this program is to map the names in RA to names in publically available datasets like EPA's eGrid
alorithm could have other applications for EIA or even SNL utility data
'''
df = pd.read_excel(r'C:\Users\kgleeson\Documents\Python Scripts\Egrid_RA_Utility_Mapping.xlsx', "RA")
RA_utilities = df['RA Utility'].tolist()
df = pd.read_excel(r'C:\Users\kgleeson\Documents\Python Scripts\Egrid_RA_Utility_Mapping.xlsx', "Egrid")
eGrid_utilities = df['Egrid Utility'].tolist()
RA_Match_List = []
Egrid_Match_List = []
Cos_Sim = []

threshold = 0.80     # if needed
match_count = 0
for key in eGrid_utilities:
    for word in RA_utilities:
        try:
            # print(key)
            # print(word)
            res = cosdis(word2vec(word), word2vec(key))
            # print(res)
            # print("The cosine similarity between : {} and : {} is: {}".format(word, key, res*100))
            if res > threshold:
                # print("Found a word with cosine distance > {} : {} with original word: {}. Cosine distance = {}".format(threshold,word, key, res))
                RA_Match_List.append(word)
                Egrid_Match_List.append(key)
                Cos_Sim.append(res)
                match_count += 1
        except IndexError:
            pass
    

print(match_count, ' matches')
# print(RA_Match_List)
# print(Egrid_Match_List)

d = {'RA Utilities' : RA_Match_List, 'Egrid Utilities' : Egrid_Match_List, 'Cosine Similarity' : Cos_Sim}
df = pd.DataFrame(d, columns = ['RA Utilities','Egrid Utilities','Cosine Similarity'])

#Getting the closest match for each RA utility
g=df.loc[df.groupby('RA Utilities')['Cosine Similarity'].idxmax()]
g.to_excel('C:\\Users\\kgleeson\\Documents\\Python Scripts\\Utility_Name_Mapping_snl.xlsx')
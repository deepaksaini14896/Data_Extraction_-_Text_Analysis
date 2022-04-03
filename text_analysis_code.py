import pandas as pd
import json
import re
import copy
from nltk.tokenize import sent_tokenize, word_tokenize, SyllableTokenizer
from nltk.corpus import stopwords
import xlsxwriter

# Create word token of text
def create_word_token(article_text):
    word = re.sub('[^A-Z]',' ',article_text.upper())
    tokenize_word = word_tokenize(word)

    return tokenize_word

# Remove stop words by using stop word dictionary
def remove_stop_word_by_file(tokenize_word, stop_word):
    clean_word = list()
    
    for word in tokenize_word:
        if word not in stop_word:
            clean_word.append(word)
    
    return clean_word

# Count positive and negative score in master dictionary
def positive_negative(clean_word, master_dict):
    positive_score = 0
    negative_score = 0

    positive_word = master_dict.query('Positive != 0')
    negative_word = master_dict.query('Negative != 0')
    
    positive_word = positive_word['Word'].tolist()
    negative_word = negative_word['Word'].tolist()
 
    for word in clean_word:
        if word in positive_word:
            positive_score += 1
        elif word in negative_word:
            negative_score -= 1
    
    negative_score *= -1
    
    return  positive_score, negative_score

# Find polarity score using positive and negative score
def polarity(positive_score, negative_score):
    polarity_score = (positive_score-negative_score)/((positive_score+negative_score)+0.000001)
    
    return polarity_score

# Find subjectivity score using clean word with the help of positive and negative score
def subjectivity(positive_score, negative_score, clean_word):
    total_clean_word = len(clean_word)
    subjectivity_score = (positive_score+negative_score)/((total_clean_word)+0.000001)
    
    return subjectivity_score
    
# Count complex word those have more than two syllable using tokenize word
def complex_word(tokenize_word):
    count = 0
    
    for word in tokenize_word:
        st = SyllableTokenizer()
        syllable = st.tokenize(word.lower())
        
        if len(syllable) > 2:
            count += 1
    
    return count
         
# Remove stop word using nltk(natural language toolkit)
def remove_stop_word_by_nltk(tokenize_word):
    stop_words = set(stopwords.words('english'))
    count = 0
    
    for word in tokenize_word:
        if word.lower() not in stop_words:
            count += 1
    
    return count

# Count syllable
def syllable_count(tokenize_word):
    vowels_count = 0
    word_count = len(tokenize_word)
    vowels = "AEIOU"
    
    for word in tokenize_word:
        if len(word) > 1 and word[-2:] in ["ES", "ED"]:
            vowels_count -= 1

        for c in word:
            if c in vowels:
                vowels_count += 1
    
    syllable_word_per_count = vowels_count/word_count
    
    return syllable_word_per_count

# Count personal pronoun count with the help of re(regular expression)
def personal_pronouns_count(article_text):
    personal_pronouns_word = re.findall(r"\bI\b|\bi\b|\bWe\b|\bwe\b|\bMy\b|\bmy\b|\bOurs\b|\bours\b|\bus\b", article_text)
    personal_pronouns = len(personal_pronouns_word)
    
    return personal_pronouns

# Find average word length
def average_word(tokenize_word):
    total_number_of_word = len(tokenize_word)
    total_number_of_character = 0
    
    for word in tokenize_word:
        total_number_of_character += len(word)
    
    average_word_length = total_number_of_character/total_number_of_word
    
    return average_word_length 

# Read input file
input_file = pd.read_excel("Raw_Data/Input.xlsx")

# Read master dictionary
master_dict = pd.read_csv("Raw_Data/Loughran-McDonald_MasterDictionary_1993-2021.csv")

# Read stop word dictionary
stop_word = pd.read_csv("Raw_Data/StopWords_Generic.txt")

# Create output file 
workbook = xlsxwriter.Workbook('Text_Analysis/Output.xlsx')

# Add sheet in workbook
worksheet = workbook.add_worksheet()

# Create space in column
worksheet.set_column(0, 14, 20)

# Create column name
column_name = ["URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE",
               "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE",
              "COMPLEX WORD COUNT", "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"]

# Add bold format in column 
bold = workbook.add_format({'bold': True})

# Write column name
worksheet.write_row(0, 0, column_name, bold)


#for i in range(len(input_file)):
for i in range(0,30):
    url_id = int(input_file.iloc[i,0])
    
    url = input_file.iloc[i,1]
    
    file_object = open("Data_Extraction/" + str(url_id) + ".txt", "r")
    
    file_data = file_object.read()
    
    file_dic = json.loads(file_data)
    
    article_text = file_dic["Article_Text"]
    
    tokenize_word = create_word_token(article_text)
    
    number_word = len(tokenize_word)
    
    number_sentence = len(sent_tokenize(article_text))
    
    clean_word = remove_stop_word_by_file(tokenize_word, stop_word)
  
    positive_score, negative_score = positive_negative(clean_word, master_dict)
    
    polarity_score = polarity(positive_score, negative_score)
    
    subjectivity_score = subjectivity(positive_score, negative_score, clean_word)    
    
    avg_sentence_length = number_word/number_sentence
    
    complex_word_count = complex_word(tokenize_word)
    
    percentage_of_complex_words = 100*(complex_word_count/number_word)

    fog_index = 0.4*(avg_sentence_length+percentage_of_complex_words)
    
    avg_number_of_words_per_sentence = number_word/number_sentence

    word_count = remove_stop_word_by_nltk(tokenize_word)
    
    syllable_per_word = syllable_count(tokenize_word)
    
    personal_pronouns = personal_pronouns_count(article_text)
    
    avg_word_length = average_word(tokenize_word)
    
    worksheet.write_row(i+1, 0, [url_id, url, positive_score, negative_score, polarity_score, subjectivity_score,
                               avg_sentence_length, percentage_of_complex_words, fog_index, avg_number_of_words_per_sentence,
                               complex_word_count, word_count, syllable_per_word, personal_pronouns, avg_word_length])
    
# Close workbook
workbook.close()

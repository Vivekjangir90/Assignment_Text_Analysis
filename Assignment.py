from bs4 import BeautifulSoup
import pandas as pd
import requests
from textblob import TextBlob
import re
import nltk
from nltk.tokenize import word_tokenize
import nltk
# nltk.download('punkt') #Downloading package PUNKT
# nltk.download('cmudict') #Downloading package cmudict

df_urls = pd.read_excel("Input.xlsx")
# print(df_urls["URL"])
for num,url_index in enumerate(df_urls["URL"]):
    # print(url_index)

    if num >= 0:
        while True:
            try:
                data = requests.get(url_index)
                # If the request is successful, break out of the loop
                break
            except Exception as e:
                # Handle the exception gracefully
                print(f"An error occurred while making the request: {e}")

        data = BeautifulSoup(data.content, "html.parser")
        # print(data.title)

        str_div = str(data(class_="td-post-content tagdiv-type"))
        pera = BeautifulSoup(str_div, "html.parser")
        pera = str(pera.text)
        pera = pera.replace("[", "").replace("]", "")


        # Tokenize the text into sentences
        total_sentences = nltk.sent_tokenize(pera)
        # Tokenize the text into words
        total_words_before_clean = nltk.word_tokenize(pera)

        avg_word_sen_list=[]
        for i in total_sentences:
            avg_word_sen = nltk.word_tokenize(i)
            avg_word_sen_list.append(len(avg_word_sen))

        # Remove special characters using regular expression
        pera = re.sub(r'[^\w\s]', '', pera)

        characters_counts = len(pera) 
        # Split the string into words
        pera = pera.split()
        negative = []
        positive = []

        neg = open("negative.txt", "w")
        pos = open("positive.txt", "w")

        stop_Auditor = open("StopWords\StopWords_Auditor.txt", "r")
        stop_Currencies = open("StopWords\StopWords_Currencies.txt", "r")
        stop_DatesandNumbers = open("StopWords\StopWords_DatesandNumbers.txt", "r")
        stop_Generic = open("StopWords\StopWords_Generic.txt", "r")
        stop_GenericLong = open("StopWords\StopWords_GenericLong.txt", "r")
        stop_Geographic = open("StopWords\StopWords_Geographic.txt", "r")
        stop_Names = open("StopWords\StopWords_Names.txt", "r")

        # neg_r_content = neg_r.read().split()
        # pos_r_content = pos_r.read().split()
        stop_Auditor_content = stop_Auditor.read().split()
        stop_Currencies_content = stop_Currencies.read().split()
        stop_DatesandNumbers_content = stop_DatesandNumbers.read().split()
        stop_Generic_content = stop_Generic.read().split()
        stop_GenericLong_content = stop_GenericLong.read().split()
        stop_Geographic_content = stop_Geographic.read().split()
        stop_Names_content = stop_Names.read().split()

        for i in pera:
            blob = TextBlob(i)
            pol = blob.sentiment.polarity

            transformed_versions = [i.lower(), i.upper(), i.capitalize()]

            if all(transformed_version not in stop_list for transformed_version in transformed_versions for stop_list in [positive, negative, stop_Auditor_content, stop_Currencies_content, stop_DatesandNumbers_content, stop_Generic_content, stop_GenericLong_content, stop_Geographic_content, stop_Names_content]):

                if pol < 0 and i not in negative :
                    neg.write(i + '\n')  # Write the sentence to the negative file
                    negative.append(i)

                else:
                    if i not in positive:
                        pos.write(i + '\n')  # Write the sentence to the positive file
                        positive.append(i)
            # else:
                # print(i)
        pos.close()
        neg.close()

        positive_score = 0
        negative_score = 0

        neg_r = open("negative.txt", "r")
        pos_r = open("positive.txt", "r")
        neg_word = neg_r.read()
        pos_word = pos_r.read()


        neg_word_tokenizer = word_tokenize(neg_word)
        neg_sys_word = open("MasterDictionary/negative-words.txt", "r")
        neg_sys_word = neg_sys_word.read()
        neg_sys_word = word_tokenize(neg_sys_word)

        for i in neg_word_tokenizer:
            if i in neg_sys_word:
                negative_score +=1
        print(negative_score)

        pos_word_tokenizer = word_tokenize(pos_word)
        pos_sys_word = open("MasterDictionary/positive-words.txt", "r")
        pos_sys_word = pos_sys_word.read()
        pos_sys_word = word_tokenize(pos_sys_word)

        for i in pos_word_tokenizer:
            if i in pos_sys_word:
                positive_score +=1
        print(positive_score)

        vowels = 'aeiouAEIOU'  
        vowel_count = 0  
        for char in neg_word: 
            if char in vowels:
                vowel_count += 1
        for char in pos_word:
            if char in vowels:
                vowel_count += 1

        # syllables_vowels =   'ae' or 'ai' or 'ao' or 'au' or 'ea' or 'ei' or 'eo' or 'eu' or 'ia' or 'ie' or 'io' or 'iu' or 'oa' or 'oe' or 'oi' or 'ou' or 'ua' or 'ue' or 'ui' or 'uo':
        syllables_vowels_count = 0  
        for char in neg_word: 
            if ('ae' or 'ai' or 'ao' or 'au' or 'ea' or 'ei' or 'eo' or 'eu' or 'ia' or 'ie' or 'io' or 'iu' or 'oa' or 'oe' or 'oi' or 'ou' or 'ua' or 'ue' or 'ui' or 'uo') in char:
                syllables_vowels_count += 1
        for char in pos_word:
            if ('ae' or 'ai' or 'ao' or 'au' or 'ea' or 'ei' or 'eo' or 'eu' or 'ia' or 'ie' or 'io' or 'iu' or 'oa' or 'oe' or 'oi' or 'ou' or 'ua' or 'ue' or 'ui' or 'uo') in char:
                syllables_vowels_count += 1

        negative_pronouns = ["I", "We", "my", "ours", "i", "we", "My", "Ours", "us", "Us"]
        pronouns_counts = 0
        for word in total_words_before_clean:
            if word in negative_pronouns and word != "US":
                # print(word, "is a negative word")
                pronouns_counts +=1

        Polarity_Score = ((positive_score - negative_score)/((positive_score + negative_score) + 0.00001))
        Subjectivity_Score = (positive_score + negative_score)/ ((len(neg_word_tokenizer)+len(pos_word_tokenizer)) + 0.000001)
        Average_Sentence_Length = len(total_words_before_clean) / (len(total_sentences)+0.00001)
        Percentage_of_Complex_words = (len(negative)+len(positive)) / (len(total_words_before_clean)+0.0001 )
        Fog_Index = 0.4 * (Average_Sentence_Length + Percentage_of_Complex_words)
        Complex_words_Count  = syllables_vowels_count
        Average_Number_of_Words_Per_Sentence = (len(negative)+len(positive)) / (len(total_sentences)+0.00001 )
        Words_Counts = len(positive)+len(negative)
        Syllable_Count_Per_Word = vowel_count/(len(positive) / ((len(negative)+0.0001))+0.00001)
        Average_Word_Length = characters_counts/(len(total_words_before_clean)+0.00001)



        # Read the existing Excel file into a DataFrame
        df = pd.read_excel("Output Data Structure.xlsx")

        # Insert data into the "name" column at row number 0
        row_index = num  # Row number 0 (indexing starts from 0)
        df.at[row_index, "POSITIVE SCORE"] = positive_score
        df.at[row_index, "NEGATIVE SCORE"] = negative_score
        df.at[row_index, "POLARITY SCORE"] = Polarity_Score
        df.at[row_index, "SUBJECTIVITY SCORE"] = Subjectivity_Score
        df.at[row_index, "AVG SENTENCE LENGTH"] = Average_Sentence_Length
        df.at[row_index, "PERCENTAGE OF COMPLEX WORDS"] = Percentage_of_Complex_words
        df.at[row_index, "FOG INDEX"] = Fog_Index
        df.at[row_index, "AVG NUMBER OF WORDS PER SENTENCE"] = Average_Number_of_Words_Per_Sentence
        df.at[row_index, "COMPLEX WORD COUNT"] = Complex_words_Count
        df.at[row_index, "WORD COUNT"] = Words_Counts
        df.at[row_index, "SYLLABLE PER WORD"] = Syllable_Count_Per_Word
        df.at[row_index, "PERSONAL PRONOUNS"] = pronouns_counts
        df.at[row_index, "AVG WORD LENGTH"] = Average_Word_Length

        # Write the modified DataFrame back to the Excel file
        df.to_excel("Output Data Structure.xlsx", index=False)
        # print(df["POSITIVE SCORE"])
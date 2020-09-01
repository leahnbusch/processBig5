# -*- coding: utf-8 -*-
"""
Created on Wed Aug 26 10:49:57 2020

@author: LBUSCH4
"""

import pandas as pd

#Variables

#study export has the labels and not the values. This dict maps the labels to the Likert values
replace_rankings = {"Disagree strongly" : 1, "Disagree a little" : 2, "Neutral; no opinion" : 3, "Agree a little" : 4, "Agree strongly" : 5}

inputFile = 'personality_input.xlsx'
outputFile = 'personality_output.xlsx'


#these variables make it easier to reference each question by name and not by number
#----------------------------------------------------------------------------------
q1 = 'I am someone who is outgoing, sociable.'
q2 = 'I am someone who is compassionate, had a soft heart.'
q3 = 'I am someone who tends to be disorganized.'
q4 ='I am someone who is relaxed, handles stress well.'
q5 = 'I am someone who has few artistic interests.'
q6 =  'I am someone who has an assertive personality'
q7 = 'I am someone who is respectful, treats others with respect.'
q8 = 'I am someone who tends to be lazy.'
q9 = 'I am someone who stays optimistic after experiencing a setback.'
q10 = 'I am someone who is curious about many different things.'
q11 = 'I am someone who rarely feels excited or eager.'
q12 = 'I am someone who tends to find fault with others.'
q13 = 'I am someone who is dependable, steady.'
q14 = 'I am someone who is moody, has up and down mood swings.'
q15 = 'I am someone who is inventive, finds clever ways to do things.'
q16 = 'I am someone who tends to be quiet.'
q17 = 'I am someone who feels little sympathy for others.'
q18 = 'I am someone who is systematic, likes to keep things in order.'
q19 = 'I am someone who can be tense.'
q20 = 'I am someone who is fascinated by art, music, or literature.'
q21 = 'I am someone who is dominant, acts as a leader.'
q22 = 'I am someone who starts arguments with others.'
q23 = 'I am someone who has difficulty getting started on tasks.'
q24 = 'I am someone who feels secure, comfortable with self.'
q25 = 'I am someone who avoids intellectual, philosophical discussions.'
q26 = 'I am someone who is less active than other people.'
q27 = 'I am someone who has a forgiving nature.'
q28 = 'I am someone who can be somewhat careless.'
q29 = 'I am someone who is emotionally stable, not easily upset.'
q30 = 'I am someone who has little creativity.'
q31 = 'I am someone who is sometimes shy, introverted.'
q32 = 'I am someone who is helpful and unselfish with others.'
q33 = 'I am someone who keeps things neat and tidy.'
q34 = 'I am someone who worries a lot.'
q35 = 'I am someone who values art and beauty.'
q36 = 'I am someone who finds it hard to influence people.'
q37 = 'I am someone who is sometimes rude to others.'
q38 = 'I am someone who is efficient, gets things done.'
q39 = 'I am someone who often feels sad.'
q40 = 'I am someone who is complex, a deep thinker.'
q41 = 'I am someone who is full of energy.'
q42 = 'I am someone who is suspicious of others’ intentions.'
q43 = 'I am someone who is reliable, can always be counted on.'
q44 = 'I am someone who keeps their emotions under control.'
q45 = 'I am someone who has difficulty imagining things.'
q46 = 'I am someone who is talkative.'
q47 = 'I am someone who can be cold and uncaring.'
q48 = 'I am someone who leaves a mess, doesn’t clean up.'
q49 = 'I am someone who rarely feels anxious or afraid.'
q50 = 'I am someone who thinks poetry and plays are boring.'
q51 = 'I am someone who prefers to have others take charge.'
q52 = 'I am someone who is polite, courteous to others.'
q53 = 'I am someone who is persistent, works until the task is finished.'
q54 = 'I am someone who tends to feel depressed, blue.'
q55 = 'I am someone who has little interest in abstract ideas.'
q56 = 'I am someone who shows a lot of enthusiasm.'
q57 = 'I am someone who assumes the best about people.'
q58 = 'I am someone who sometimes behaves irresponsibly.'
q59 = 'I am someone who is temperamental, gets emotional easily.'
q60 = 'I am someone who is original, comes up with new ideas.'

#----------------------------------------------------------------------------------
#Functions

#some questions are asked the inverse of what we need. This inverses the values
def reverseKey(input):
    if input == 5:
        return 1
    elif input == 4:
        return 2
    elif input == 3:
        return 3
    elif input == 2:
        return 4
    else:
        return 5

#Calcualate the values for O C E A and N domains and fascets separately

def O_Values(row):
    O_Itellectual_Curiosity = row[q10] + reverseKey(row[q25]) + row[q40] + reverseKey(row[q55])#10,25R, 40, 55R
    O_Aesthetic_Sensitivity = reverseKey(row[q5]) + row[q25] + row[q35] + reverseKey(row[q50])#5R, 20, 35, 50R
    O_Creative_Imagineation = row[q15] + reverseKey(row[q30]) + reverseKey(row[q45]) + row[q60] #Creative Imagination: 15, 30R, 45R, 60
    O = O_Itellectual_Curiosity + O_Aesthetic_Sensitivity + O_Creative_Imagineation
    return O_Itellectual_Curiosity/4, O_Aesthetic_Sensitivity/4, O_Creative_Imagineation/4, O/12

def C_Values(row):
    C_Organization = reverseKey(row[q3]) + row[q18] + row[q33] + reverseKey(row[q48])#3R, 18, 33, 48R
    C_Productiveness = reverseKey(row[q8]) + reverseKey(row[q23]) + row[q38] + row[q53]#Productiveness: 8R, 23R, 38, 53
    C_Responsibility = row[q13] + reverseKey(row[q28]) + row[q43] + reverseKey(row[q58]) #13, 28R, 43, 58R
    C = C_Organization + C_Productiveness + C_Responsibility
    return C_Organization/4, C_Productiveness/4, C_Responsibility/4, C/12

def E_Values(row):
    E_Sociabilty = row[q1] + reverseKey(row[q16]) + reverseKey(row[q31]) + row[q46] #Sociability: 1, 16R, 31R, 46
    E_Assertiveness = row[q6] + row[q21] + reverseKey(row[q36]) + reverseKey(row[q51]) #Assertiveness: 6, 21, 36R, 51R
    E_Energy_Level = reverseKey(row[q11]) + reverseKey(row[q26]) + row[q41] + row[q56] #Energy Level: 11R, 26R, 41, 56
    E = E_Sociabilty + E_Assertiveness + E_Energy_Level
    return E_Sociabilty/4, E_Assertiveness/4, E_Energy_Level/4, E/12

def A_Values(row):
    A_Compassion = row[q2] + reverseKey(row[q17]) + row[q32] + reverseKey(row[q47]) #Compassion: 2, 17R, 32, 47R
    A_Respectfulness = row[q7] + reverseKey(row[q22]) + reverseKey(row[q37]) + row[q52]#Respectfulness: 7, 22R, 37R, 52
    A_Trust = reverseKey(row[q12]) + row[q27] + reverseKey(row[q42]) + row[q57] #Trust: 12R, 27, 42R, 57
    A = A_Compassion + A_Respectfulness + A_Trust
    return A_Compassion/4, A_Respectfulness/4, A_Trust/4, A/12

def N_Values(row):
    N_Anxiety = reverseKey(row[q4]) + row[q19] + row[q34] + reverseKey(row[q49]) #Anxiety: 4R, 19, 34, 49R
    N_Depression = reverseKey(row[q9]) + reverseKey(row[q24]) + row[q39] + row[q54] #Depression: 9R, 24R, 39, 54
    N_Emotional_Volatility = row[q14] + reverseKey(row[q29]) + reverseKey(row[q44]) + row[q59]#Emotional Volatility: 14, 29R, 44R, 59
    N = N_Anxiety + N_Depression + N_Emotional_Volatility
    return N_Anxiety/4, N_Depression/4, N_Emotional_Volatility/4, N /12

#----------------------------------------------------------------------------------
#begin main

Personality = pd.read_excel(inputFile)
Personality = Personality.replace(replace_rankings)

#replace all nan's with 3
Personality = Personality.fillna(value=3)

column_names = ["username", "O","C", "E", "A", "N", "O_Itellectual_Curiosity", "O_Aesthetic_Sensitivity", "O_Creative_Imagineation", \
                "C_Organization", "C_Productiveness", "C_Responsibility", \
                "E_Sociabilty", "E_Assertiveness", "E_Energy_Level", \
                "A_Compassion", "A_Respectfulness", "A_Trust", \
                "N_Anxiety", "N_Depression", "N_Emotional_Volatility"]
processedPersonalityResults = pd.DataFrame( columns = column_names)

for index, row in Personality.iterrows():
   O_Itellectual_Curiosity, O_Aesthetic_Sensitivity, O_Creative_Imagineation, O = O_Values(row)
   C_Organization, C_Productiveness, C_Responsibility, C = C_Values(row)
   E_Sociabilty, E_Assertiveness, E_Energy_Level, E = E_Values(row)
   A_Compassion, A_Respectfulness, A_Trust, A = A_Values(row)
   N_Anxiety, N_Depression, N_Emotional_Volatility, N = N_Values(row)

   processedPersonalityResults = processedPersonalityResults.append({
        "username": row["Created By"],
        "O" : O,"C" : C, "E" : E , "A" : A, "N" : N, "O_Itellectual_Curiosity" : O_Itellectual_Curiosity ,\
		"O_Aesthetic_Sensitivity" : O_Aesthetic_Sensitivity, "O_Creative_Imagineation" : O_Creative_Imagineation, \
		"C_Organization" : C_Organization, "C_Productiveness" : C_Productiveness, "C_Responsibility" : C_Responsibility, \
		"E_Sociabilty" : E_Sociabilty, "E_Assertiveness" : E_Assertiveness, "E_Energy_Level" : E_Energy_Level, \
		"A_Compassion" : A_Compassion, "A_Respectfulness" : A_Respectfulness, "A_Trust" : A_Trust, \
		"N_Anxiety" : N_Anxiety, "N_Depression" : N_Depression, "N_Emotional_Volatility" : N_Emotional_Volatility
    }, ignore_index=True)

processedPersonalityResults.to_excel(outputFile)

#FYI
#(http://www.colby.edu/psych/personality-lab/).


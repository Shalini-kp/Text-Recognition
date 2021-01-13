#!/usr/bin/python
import os
import re
import json
import xlsxwriter

import string
import difflib
import datetime
from statistics import mean

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('voiceof29_test.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})
formatk = workbook.add_format({'num_format': 'ss'})

# number of columns and their names
alphabets = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 
'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']

column_name = ['Names', 'Base Word count', 'Recorded Word count', 'Word Difference', 'Word Inserted',
'Word Dropped', 'Word Omitted', 'No of Word Substituted by', 'Total Pauses between words', 'Pauses between words per minutes',
'Pauses per 100 words', 'Pauses less than 200ms', 'Cumulative value of pauses(in secs)', 'words per minute',
'Time taken for 100 words(in secs)', 'Total time taken(in secs)', 'Repetitions per 100 words', 'Repetitions per minutes',
'Total Repetition', 'No. of fillers in Sentence', 'Overall Confidence', 'Confidence excluding substituted words',
'Confidence Including substituted words(>0.5) ', 'Confidence Including substituted words(<0.5) ']

# set the columns ready 
for col_value, col_name in enumerate(column_name):
    column_aplha = alphabets[col_value]
    worksheet.set_column((column_aplha + ':' + column_aplha), (len(col_name) + 10)) # Widen the first column to make the text clearer.
    worksheet.write((column_aplha + '1'), col_name, bold) # writing column name

filePath = "/Users/tb/Desktop/RandD/TextRecognition/Resources/Audio_json/"
user_collection = os.listdir(filePath)
userCollection = filter(lambda x: re.search(".json", x), user_collection)

# define class to hold the items of voice recorded
class VoiceDetails:
    def __init__(self, startTime, endTime, content, confidence):
        self.start_time, self.end_time, self.content, self.confidence = startTime, endTime, content, confidence

    def updateConfidence(self, confidence):
        self.confidence = confidence

# remove special characters
def remove_special_charaters(textFile):
    text_File = []
    for i in textFile:
        b = ''.join(' ' if c.isalpha() == False and c != '\'' else c for c in i)
        if len(b.split()) > 1:
            for j in b.split():
                text_File.append(j.strip())
        else: 
            text_File.append(b.strip())
    return text_File

# find out repitations 
def repitating_words(wordlist):
    repItem = []
    for index, i in enumerate(wordlist):
        if index > 0:
            if wordlist[index - 1] == i:
                if i not in repItem:
                    repItem.append(i)

            elif len(wordlist) - 1 > index:
                if wordlist[index + 1] == i:
                    if i not in repItem:
                        repItem.append(i)
    return repItem

# confidence
def calculate_mean_confidence(confidenceValue):
    if confidenceValue is not None and len(confidenceValue) > 0:
        confidenceValue = mean(confidenceValue)
    return round(confidenceValue, 2)
  
if __name__ == "__main__":
    for rowCount, userName in enumerate(userCollection):

        rowCount += 2
        rowCount = str(rowCount)
        isFirst = True
        bM_Pause = 2.0
        bM_Speed = 0.2
        totalTimeTaken, fillersCount = 0.0, 0.0
        Voice_Details, Voice_Details_ExSub, Voice_Details_InM5, Voice_Details_InL5  = [], [], [], []
        TotalPauses, TotalSpeed, cumulativePause = [], [] , []
        substituted, substitutedBy, inserted, dropped = [], [], [], []

        # fetch next values
        def obtain_next_value(i, nextIndex, totalLength):

            if nextIndex < totalLength:
                    j = item[nextIndex]
                    if j.get("start_time") is not None:
                        calculated_pause = float(j["start_time"]) - float(i["end_time"])
                        cumulativePause.append(calculated_pause)

                        if calculated_pause >= bM_Pause:
                            TotalPauses.append(calculated_pause)
                        elif calculated_pause <= bM_Speed:
                            TotalSpeed.append(calculated_pause)
                    else:
                        obtain_next_value(i, nextIndex + 1, totalLength)

        # substituted and substituted by
        def substitue(currentElement, value):
            if currentElement == '+':
                substitutedBy.append(value)
            elif currentElement == '-':
                substituted.append(value)

        # Inserted and dropped
        def insert_drop(element, value):
            if (element == '+'):
                inserted.append(value)
            elif (element == '-'):
                dropped.append(value)

        # open and read json
        openFile = open(filePath + userName,'r')
        jsonDict = json.load(openFile)

        baseFile = "He does not expose himself needlessly to danger, since there are few things for which he cares sufficiently; but he is willing, in great crises, to give even his life–knowing that under certain conditions it is not worthwhile to live. He is of a disposition to do men service, though he is ashamed to have a service done to him. To confer a kindness is a mark of superiority; to receive one is a mark of subordination… He does not take part in public displays… He is open in his dislikes and preferences; he talks and acts frankly, because of his contempt for men and things… He is never fired with admiration, since there is nothing great in his eyes. He cannot live in complaisance with others, except it be a friend; complaisance is the characteristic of a slave… He never feels malice, and always forgets and passes over injuries… He is not fond of talking… It is no concern of his that he should be praised, or that others should be blamed. He does not speak evil of others, even of his enemies, unless it be to themselves. His carriage is sedate, his voice deep, his speech measured; he is not given to hurry, for he is concerned about only a few things; he is not prone to vehemence, for he thinks nothing very important. A shrill voice and hasty steps come to a man through care… He bears the accidents of life with dignity and grace, making the best of his circumstances, like a skilful general who marshals his limited forces with the strategy of war… He is his own best friend, and takes delight in privacy whereas the man of no virtue or ability is his own worst enemy, and is afraid of solitude."
        recordedFile = jsonDict["results"]["transcripts"][0]["transcript"]

        # split the words, by default through the space
        baseFile = baseFile.split()
        recordedFile = recordedFile.split()

        baseFile = remove_special_charaters(baseFile)
        recordedFile = remove_special_charaters(recordedFile)

        baseFile = [x for x in baseFile if x is not ""] 
        recordedFile = [x for x in recordedFile if x is not ""]

        # word count
        base_wordcount = len(baseFile)
        recorded_wordcount = len(recordedFile)

        item = jsonDict["results"]["items"]
        item = [x for x in item if x is not None] 

        for index,i in enumerate(item):
            if i.get("start_time") is not None and i.get("end_time") is not None:

                if isFirst == True:
                    isFirst = False
                    first_startTime = float(i["start_time"])
                    cumulativePause.append(first_startTime)

                    # starting pause lag
                    if first_startTime >= bM_Pause:
                        TotalPauses.append(first_startTime)
                    
                # calulate pauses
                obtain_next_value(i, index + 1, len(item))  
        
                # words to be meaured
                content_words = i["alternatives"][0]["content"]
                confidence_value = i["alternatives"][0]["confidence"]

                # create object of class VoiceDetails
                Voice_Details.append(VoiceDetails(i["start_time"], i["end_time"], content_words, confidence_value))
                Voice_Details_ExSub.append(VoiceDetails(i["start_time"], i["end_time"], content_words, confidence_value))
                Voice_Details_InM5.append(VoiceDetails(i["start_time"], i["end_time"], content_words, confidence_value))
                Voice_Details_InL5.append(VoiceDetails(i["start_time"], i["end_time"], content_words, confidence_value))

                # Total time taken
                totalTimeTaken = float(i["end_time"])

        wordsList = [x.content for x in Voice_Details]
        wordsList = [x for x in wordsList if x is not None]

        # total pauses between words
        TotalPauses = [x for x in TotalPauses if x is not None]

        # total speed between words
        TotalSpeed = [x for x in TotalSpeed if x is not None]

        # pauses per 100 words
        pausePer100Word = round(((len(TotalPauses) / len(wordsList)) * 100), 2)

        # pauses per minute
        pausePerMinute = round(((len(TotalPauses) / totalTimeTaken) * 60), 2)

        # words per minute
        wordsPerMinute = round(((len(wordsList) / totalTimeTaken) * 60), 2)

        # Time taken for 100 words and total time taken
        TimeTaken100Words = round(((totalTimeTaken / len(wordsList)) * 100), 2)

        # No. of fillers
        for index,i in enumerate(recordedFile):
            if i in ["huh", "uh", "erm", "um", "hmm", "ah"]:
                fillersCount += 1    

        # compare and list out the difference
        d = difflib.Differ()     
        diff = list(d.compare(baseFile, recordedFile))
        [diff.remove(line) for line in diff if line[0] == '?']
        #print('\n'.join(diff))

        # inserted, dropped, substituted and substituted by
        for index,i in enumerate(diff):
            if index > 0:
                currentElement = diff[index][0]
                previousElement = diff[index - 1][0]

                if (previousElement == '+' or previousElement == '-') and (currentElement == '+' or currentElement == '-'):
                    substitue(currentElement, diff[index][1:])
                elif len(diff) - 1 > index:
                    nextElement = diff[index + 1][0]
                    if (nextElement == '+' or nextElement == '-') and (currentElement == '+' or currentElement == '-'):
                        substitue(currentElement, diff[index][1:])
                    else:
                        insert_drop(diff[index][0], diff[index][1:])
                else:
                    insert_drop(diff[index][0], diff[index][1:])

        # repitation of words
        totalRepitation = repitating_words(wordsList)
        repitation100Words = round(((len(totalRepitation)/len(wordsList)) * 100), 2)
        repitationPerMinute = round(((len(totalRepitation)/totalTimeTaken) * 60), 2)

        # overall confidence
        added_words = [line for line in diff if line[0] != "-"]
        for index, it in enumerate(Voice_Details):

                # confidence excluding substituted words
                if added_words[index][0] == "+":
                    Voice_Details_ExSub[index].confidence = 0.0

                    # confidence including substituted words(if it is more than 50%)
                    moreThan5Confidence = float(Voice_Details_InM5[index].confidence)
                    if moreThan5Confidence > 0.5:
                        Voice_Details_InM5[index].confidence = moreThan5Confidence / 2

                        # confidence including substituted words(if it is less than 50%)
                        Voice_Details_InL5[index].confidence = 0.0

        overallConfidence = calculate_mean_confidence([float(x.confidence) for x in Voice_Details])
        excludeSub = calculate_mean_confidence([float(x.confidence) for x in Voice_Details_ExSub])
        includeM5Sub = calculate_mean_confidence([float(x.confidence) for x in Voice_Details_InM5])
        includeL5Sub = calculate_mean_confidence([float(x.confidence) for x in Voice_Details_InL5])

        cummulative_pause = str(datetime.timedelta(seconds=sum(cumulativePause)).seconds)
        time_taken_100_words = str(datetime.timedelta(seconds=TimeTaken100Words).seconds)
        total_timetaken = str(datetime.timedelta(seconds=totalTimeTaken).seconds)

        # write to file
        calculted_values = [userName.split('.')[0], base_wordcount, recorded_wordcount, abs(base_wordcount - recorded_wordcount), 
        len(inserted), len(dropped), len(substituted), len(substitutedBy), len(TotalPauses), pausePerMinute, pausePer100Word,
        round(sum(TotalSpeed), 2), cummulative_pause + '_format_', wordsPerMinute, time_taken_100_words + '_format_', total_timetaken + '_format_',
        repitation100Words, repitationPerMinute, len(totalRepitation), fillersCount, overallConfidence, excludeSub, includeM5Sub, includeL5Sub]

        for col_value, cal_value in enumerate(calculted_values):
            if re.search("_format_", str(cal_value)) is None:
                worksheet.write(alphabets[col_value] + rowCount, cal_value)
            else:
                worksheet.write(alphabets[col_value] + rowCount, cal_value.strip("_format_"), formatk)

workbook.close()
print("Done")
#pending stop words
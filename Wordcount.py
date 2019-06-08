from collections import *
import matplotlib.pyplot as plt
from matplotlib import cm
import numpy as np
import operator
import re
import os
import win32com.client

stop_words = open("Common Words.txt", 'r').read()
common_words = stop_words.replace(" ", "|")
removable_words = "(\s+)(" + common_words + ")(\s+)"

print stop_words


#converts file into txt
#DOC_FILEPATH = "path\doc.docx" #raw_input("Folder Path ")
#doc = win32com.client.GetObject(DOC_FILEPATH)
#text = doc.Range().Text

text = open("twitter.txt", 'r').read()

with open("something.txt", "wb") as f:
    f.write(text.encode("utf-8"))

#gets txt file and does something with it
folder = "something.txt" 

top_blank = raw_input("How Many Words? ")
articles = raw_input("Keep Common Words? (Y/N) ")
keep_other = raw_input("Show Other Category? (Y/N) ")

top_ten = []  #usage of top ten words
the_rest = [] #collection of usage of every other word

legend = []   #10 most used words

#def article_remover(dirty_text):
#    articlefree_text = re.sub('(\s+)(a|an|and|the)(\s+)','\1\3', dirty_text)
#    return articlefree_text

def word_counter(to_be_counted):
    if articles.lower() == 'y':
        f = open(to_be_counted, 'r')                             #opens file C:\users\mattd\Desktop\GEaH.txt
        text = f.read()                                          #file to string 
        clean_text = re.sub(r'[?|$|.|!|\|:|/|(|)|,|"]',r'',text)    #cleans text of all symbols 
        low_clean_text = clean_text.lower()                      #sets all words to lower case
        words = low_clean_text.split()                           #string to array
    else:
        f = open(to_be_counted, 'r')                                   #opens file C:\users\mattd\Desktop\GEaH.txt
        text = f.read()                                                #file to string 
        clean_text = re.sub(r'[?|$|.|!|\|/|:|(|)|,"]',r'',text)          #cleans text of all symbols 
        articlefree_text = re.sub(removable_words,'\1\3', clean_text)  #removes stop words
        low_clean_text = articlefree_text.lower()                      #sets all words to lower case
        words = low_clean_text.split()                                 #string to array

    top_ten_counted = OrderedDict(Counter(words).most_common(int(top_blank))) #dictionary of top ten

    print "#####################################"
    print top_ten_counted
    print "#####################################"
      
    the_rest_counted = dict(Counter(words))                  #dictionary of everything

    for key, value in top_ten_counted.iteritems() :
        top_ten.append(value)
        legend.append(key)
        print key, value

    for key, value in the_rest_counted.iteritems():
        the_rest.append(value)
 

    print sum(the_rest)
    print sum(top_ten)
    print (sum(the_rest) - sum(top_ten))
   
    if keep_other.lower() == 'y':
        other = (sum(the_rest) - sum(top_ten)) #the number of all words minus top ten
        top_ten.append(other)
        legend.append('"Other"')


    cs = cm.Set1(np.arange(40)/40)

    #Pie Graph
    plt.figure(1)
    explode = [0] * (int(top_blank))
    explode.append(.25)
    
    plt.axis('equal')
    plt.title("Percent Usage of Words")
    p = plt.pie(top_ten, labels=legend, shadow = True, startangle = 90, autopct='%1.1f%%')

    #Bar Graph
    plt.figure(2)
    objects = legend[:len(legend)-1]       #words

    y_pos = np.arange(len(objects))
    plt.yticks(y_pos, objects)             #y-axis
    performance = top_ten[:len(top_ten)-1] #x-axis

    plt.title("Word Count")
    plt.xlabel("Times Used")
    plt.ylabel("Words")
    b = plt.barh(y_pos,performance)

    plt.gca().xaxis.grid(True)
    print objects


    return b,p
word_counter(folder)

plt.show()

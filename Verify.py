import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
#read the xlsx file
data = pd.read_excel("Termes-interdits-poclain-CN-YUAN-V2.xlsx")
first_line = pd.read_excel("Termes-interdits-poclain-CN-YUAN-V2.xlsx", nrows=1,header=None)
#the first column contains the terms of each website, now we need to store them in a list
websites = data.iloc[:,0].values.tolist()
#the items in first_line from 8th column are the forbidden terms, now we need to store them in a list
forbidden_terms = first_line.iloc[0,7:].values.tolist()

print(forbidden_terms)

#now we create a double-key dictionary named "Output" with websites as the first key and forbidden terms as the second key
#the value of the dictionary is the number of times the forbidden term appears in the website, which is initialized to 0
#the dictionary is like this:{website1:{forbidden_term1:0,forbidden_term2:0,...},website2:{forbidden_term1:0,forbidden_term2:0,...},...}
Output = {}
for website in websites:
    Output[website] = {}
    for forbidden_term in forbidden_terms:
        Output[website][forbidden_term] = 0

#now we need to download the html of each website and store the html files in ./webs
#before that, we need to clear the ./webs folder
if os.path.exists("./webs"):
    for file in os.listdir("./webs"):
        os.remove("./webs/"+file)
else:
    os.mkdir("./webs")

#test if the websites are valid
#write them into a txt file
with open("valid_websites.txt","w") as f:
    for website in websites:
        f.write(website+"\n")

#download the html of each website and store them in ./webs and set the timeout to 10s
#write the exception message into log.txt
with open("log.txt","w") as f:
    for website in websites:
        try:
            r = requests.get(website,timeout=10)
            with open("./webs/"+website.split("/")[-1]+".html","w",encoding="utf-8") as f1:
                f1.write(r.text)
        except Exception as e:
            f.write(website+": "+str(e)+"\n")

#close the log.txt
f.close()
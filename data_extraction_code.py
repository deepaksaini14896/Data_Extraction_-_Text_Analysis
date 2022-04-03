import requests
import json
from bs4 import BeautifulSoup
import pandas as pd

#Read input file into dataframe object.
df = pd.read_excel("Raw_Data/Input.xlsx")

for i in range(len(df)):
    url_id = int(df.iloc[i,0])
    
    url_link = df.iloc[i,1]
    
    user_agent = { "User-Agent" : "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36" }
    
    rqt = requests.get(url_link, headers = user_agent)
    
    soup = BeautifulSoup(rqt.content, 'html5lib')
    
    article_title = soup.find('h1', attrs = { "class" : "entry-title" }).text
    
    article=soup.find('div', attrs = { "class" : "td-post-content" })

    #Handle exception when pre tag not found in article;
    try:
    	#Remove pre tag with the help of extract function
        article.find("pre").extract().text
    
    except AttributeError:
        pass
    
    article_text = article.text
    
    #Create a dictionary to collect data.
    article_dic = {
                    	"URL_ID" : url_id,
                    	"Article_Title" : article_title,
                    	"Article_Text" : article_text.rstrip().lstrip()
                    }
    
    #Convert dictionary to json format
    article_data = json.dumps(article_dic ,indent=4)
    
    #Create file in Data Extraction folder with the name of url id's.
    with open("Data_Extraction/" + str(url_id) + ".txt", "w") as outfile:
        outfile.write(article_data)
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
from urllib.parse import quote
import re
import requests
import sys
import os
import time
import json
import html2text
import pandas as pd
h = html2text.HTML2Text()
h.ignore_links = True
def parse_description(description_tag):
    description_text = h.handle(str(description_tag)) 
    return description_text


if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
    
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
 
os.chdir(application_path)

google_session_profile = os.path.join(application_path,"google_profile")

headers_html = {
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.81 Safari/537.36 Edg/104.0.1293.54",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-User": "?1",
    "Sec-Fetch-Dest": "document",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "en-US,en;q=0.9",
    'referer':'https://www.google.com/'
}

def initialize_chrome(_from="facebook",retry=0):
    global driver 
    try:
        print("Initializing chromedriver.")
        options = Options()
        options.add_argument(f'--user-data-dir={google_session_profile}')
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument('--log-level=3')
        #options.add_argument('--blink-settings=imagesEnabled=false')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
        time.sleep(3)
        return True
    except Exception as e:
        print(e)
        pass
        
        
def google_reviews(url):

    biz_id = re.search(r"1s(0.*?\:.*?)[^a-zA-Z\d\s:]",url)
    if not biz_id:
        driver.quit()
        print("Not a valid url.")
        sys.exit(1)
    
    biz_id = biz_id.groups()[0]
    
    if not initialize_chrome():
        print("Error : ", "Failed to start chromedriver, Retry!")
        sys.exit(1)
      
    driver.get(url)
    
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="bJzME Hu9e2e tTVLSc"]')))
    except:
        print("Error : ", "Couldn't find reviews section, Retry!")
        sys.exit(1)
    
    soup = BeautifulSoup(driver.page_source,'html.parser')
    rating_and_review_div = soup.find('div',{'class':'bJzME Hu9e2e tTVLSc'})
    cookies = driver.get_cookies() 
    
    if not rating_and_review_div:
        driver.quit()       
        print("Error : ","No rating and review section found on the page, Retry or check url.")
        sys.exit(1)
        
    total_reviews = rating_and_review_div.find('button',{'aria-label':re.compile("^\d+ reviews")})
    if total_reviews:
        total_reviews = total_reviews.get('aria-label','').strip() 
        total_reviews = total_reviews.removesuffix('reviews').strip()
        
    if not total_reviews:
        total_reviews = ""

    total_rating = rating_and_review_div.find('span',{'aria-label':re.compile("^\s+?\d+\.\d+ stars")})
    if total_rating:
        total_rating = total_rating.get('aria-label','').strip() 
        total_rating = total_rating.removesuffix('stars').strip()
        
    if not total_rating:
        total_rating = ""  
   
    alias = rating_and_review_div.find(class_="fontHeadlineLarge").text.strip()
   
    address = rating_and_review_div.find('button',{'aria-label':re.compile("^Address\:")})
    if address:
        address = address.get('aria-label','').strip() 
        address = address.removeprefix('Address:')
        
    if not address:
        address = ""
    
    phone = rating_and_review_div.find('button',{'aria-label':re.compile("^Phone\:")})
    if phone:
        phone = phone.get('aria-label','').strip()    
        phone = phone.removeprefix('Phone:')
    if not phone:
        phone = ""
    
    opening_hours = rating_and_review_div.find('div',{'aria-label':re.compile("\w+day\, \d+[p|a]m to \d+[p|a]m")})
    if not opening_hours:
        opening_hours = rating_and_review_div.find('div',{'aria-label':re.compile("\w+day\, Open \d+ \w+")})

    if opening_hours:
        opening_hours = opening_hours.get('aria-label','')   
    if not opening_hours:
        opening_hours = ""
    
    website = rating_and_review_div.find('a',{'aria-label':re.compile("^Website\:")})
    if website:
        website = website.get('aria-label','')   
        website = website.removeprefix('Website:')
    if not website:
        website = ""
    
    
    REVIEWS = []
    next_page_token = ""
    
    headers_html['referer'] = url
    
    s = requests.Session()
    for cookie in cookies:
        s.cookies.set(cookie['name'], cookie['value'])
        
    while True:
        print("Extracted : ",len(REVIEWS))
        req = requests.Request('GET',f"https://www.google.com/async/reviewSort?yv=3&async=feature_id:{biz_id},review_source:All%20reviews,sort_by:newestFirst,is_owner:false,filter_text:,associated_topic:,next_page_token:{next_page_token},_pms:s,_fmt:json")
        prepared = s.prepare_request(req)
        prepared.headers = headers_html
        resp = s.send(prepared)      
        json_text = resp.content.decode('utf8').removeprefix(")]}'")
        
        json_data = json.loads(json_text)['localReviewsProto']
        #print(json_data)
        reviews = json_data['other_user_review']
             
        for review in reviews:
            rating = review['star_rating']['value']
            profile_pic = review['profile_photo_url']
            profile_pic = re.sub('=s(\d+)-','=s180-',profile_pic)
            reviewer_name = review['author_real_name']
            review_date = review['publish_date']['localized_date']
            if review.get('review_text'):
                review_text = parse_description(review['review_text']['full_html'])
            else:
                review_text = ""
            
            info_dict = {}
            
            info_dict['author'] = reviewer_name
            info_dict['rating'] =  rating             
            info_dict["date"] = review_date
            info_dict["text"] = review_text
            info_dict["profile_pic"] = profile_pic
            
            REVIEWS.append(info_dict)
            
                
        next_page_token = json_data.get('next_page_token','').strip()
        
        if not next_page_token:
            break
    
    basic_info = {
        "Title": alias,
        "Total Reviews": total_reviews,
        "Rating": total_rating,
        "Opening Hours": opening_hours,
        "Address": address,
        "Phone": phone,
        "Website":website
        
        
    }
    
    filename = re.sub('[\\/:*?"<>|]', '', alias)
    xlsx_file = f"{filename}.xlsx"
    writer = pd.ExcelWriter(xlsx_file, engine='xlsxwriter',engine_kwargs={'options': {'strings_to_urls': False}})
    
    basic_info_df = pd.DataFrame(basic_info,index=[0])
    basic_info_df.to_excel(writer,sheet_name=f"Info",index=False)  

    reviews_info_header = list(REVIEWS[0].keys())
    reviews_info_df = pd.DataFrame.from_records(REVIEWS,columns=reviews_info_header) 
    reviews_info_df.to_excel(writer,sheet_name=f"Reviews",index=False) 
    writer.save()
    

url = input('Enter google map business link : ').strip() 
try:
    google_reviews(url)
except Exception as e:
    pass
driver.quit()

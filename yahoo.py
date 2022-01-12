import time
import random
from random import randint
import requests
from bs4 import BeautifulSoup #A python library to help you to exract HTML information
from fake_useragent import UserAgent

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
import xlrd
import pandas as pd
df_keywords = pd.read_excel('mimiworld_keyword.xlsx', sheet_name='Sheet1', usecols="A")
workbook = xlrd.open_workbook('mimiworld_keyword.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
index=df_keywords.index
number_of_row=len(index)
 

headers = {
    'sec-ch-ua': "\" Not;A Brand\";v=\"99\", \"Google Chrome\";v=\"97\", \"Chromium\";v=\"97\"",
    'sec-ch-ua-mobile': "?0",
    'sec-ch-ua-platform': "\"Linux\"",
    'upgrade-insecure-requests': "1",
    'user-agent': "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36",
    'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    'cp-extension-installed': "Yes"
    }

proxies = {
  'http': 'http://192.186.190.73:8080',
  'https': 'http://192.186.190.73:8080'
}


data=[]
lowest1=[]
finalResults=[]
#data_name=[]
#competitor=[]

for i in range (1,number_of_row+1):
    time.sleep(random.randint(2,6))
    keyword_input=worksheet.cell(i,0).value
    print (keyword_input)

 
    # prefix="https://tw.mall.yahoo.com/search/product?disp=list&p="
    url = "https://tw.mall.yahoo.com/search/product"
     
    
    querystring = {"p":str(keyword_input),"sort":"p"}
    r = requests.request("GET", url, headers=headers, params=querystring)
    # r=requests.get(url)
    soup=BeautifulSoup(r.text,  'html.parser')
    
    #competitor_name_div=soup.findAll("div", {"class":"ListItem_shop_Z_sCW"})
     #results  of all products we get from the search
    productprice_div=soup.find("ul", class_="gridList")
    
    if productprice_div is not None:
        number_result_div=soup.find("div", {"class":"SortBar_sortBar_2CVWp SortBar_store_1U4Du"})
        total_results =number_result_div.find('span', {"class":'SortBar_sortCount_1LpL9 textEllipsis'}).text.strip().split(" ")[0].strip()
        print("Totol Results: ",total_results)
        # results.append(total_results)
        product_list =productprice_div.findAll("li",{"class":"BaseGridItem__grid___2wuJ7 imprsn BaseGridItem__multipleImage___37M7b"})
        # print(len(product_list))
        results=[]
        for i in product_list: 
            try:
                price=i.find("span",{"class":"BaseGridItem__price___31jkj"}).find('em').text.strip()
                print(price)
                prod_name =i.find("span",class_= "BaseGridItem__title___2HWui").text.strip()
                store_name =i.find("span",class_= "StoreGridItem__storeName___2dutX").text.strip()
                results.append({"keyword":keyword_input,"total_results":total_results,"price":int(price.replace("$","")),"name":prod_name, "competitor":store_name})
            except:
                price=i.find("span",{"class":"BaseGridItem__itemInfo___3E5Bx"}).find('em').text.strip()
                print(price)
                prod_name =i.find("span",class_= "BaseGridItem__title___2HWui").text.strip()
                store_name =i.find("span",class_= "StoreGridItem__storeName___2dutX").text.strip()
                results.append({"keyword":keyword_input,"total_results":total_results,"price":int(price.replace("$","")),"name":prod_name, "competitor":store_name})
        # print("1st RESULTS: ", results)
        
        prices =[]#Extract all prices  for each product in the  results above.then Get the lowest/minimum
        
        for product in results:
            prices.append(product['price'])
        
        lowest_price=min(prices) #get the  min price
        index_of_price =prices.index(lowest_price) #get index of the lowest  price      
        result_with_lowest_price =results[index_of_price]# get the product details of the product with the  lowest price from the initial search results list 
        
        print(result_with_lowest_price)
        #append the  product with the  lowest price to our final  list
        finalResults.append(result_with_lowest_price)
    
    
    else:
        
        total_results =0    
        # results.append(total_results)
        finalResults.append({"keyword":keyword_input,"total_results":total_results,"price":'N/A',"name":'N/A', "competitor":'N/A'})
        print("Totol Results: ",total_results)
      
    
    # lowest_price= [item for item in results if item['price'] ==min(results)]
    
    # print(results)
    
    # data.append({"keyword":keyword_input, "results":total_results, "prices":','.join(results), "lowest_price": results[0], "index": results.index(results[0])})


df = pd.DataFrame(finalResults)
df.to_excel('finalData.xlsx', index = False)    
    
  
    
    
         
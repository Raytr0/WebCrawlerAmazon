import json
import scrapy
from urllib.parse import urljoin
import re
# data from excel
import xlwings as xw
 
# Specifying a sheet
ws = xw.Book("Book2.xlsx").sheets['Sheet1']
class AmazontestSpider(scrapy.Spider):
    name = "amazonTest"
    global requestedPages
    requestedPages = ws.range("B3").value ## Number of pages
    global inOut
    inOut = ws.range("B6").value
    global excludeBrands
    excludeBrands = [ws.range("B8").value]
    def start_requests(self):
        keyword_list = [ws.range("B2").value]
        for keyword in keyword_list:
            amazon_search_url = f'https://www.amazon.com/s?k={keyword}&page=1'
            yield scrapy.Request(url=amazon_search_url, callback=self.discover_product_urls, meta={'keyword': keyword, 'page': 1})

    def discover_product_urls(self, response):
        page = response.meta['page']
        keyword = response.meta['keyword'] 

        ## Discover Product URLs
        search_products = response.css("div.s-result-item[data-component-type=s-search-result]")
        for product in search_products:
            relative_url = product.css("h2>a::attr(href)").get()
            product_url = urljoin('https://www.amazon.com/', relative_url).split("?")[0]
            yield scrapy.Request(url=product_url, callback=self.parse_product_data, meta={'keyword': keyword, 'page': page})
            
        ## Get All Pages
        if page == 1:
            available_pages = response.xpath(
                '//*[contains(@class, "s-pagination-item")][not(has-class("s-pagination-separator"))]/text()'
            ).getall()

        if available_pages < requestedPages:
            last_page = available_pages[-1]
        else:
            last_page = requestedPages

            for page_num in range(2, int(last_page)):
                amazon_search_url = f'https://www.amazon.com/s?k={keyword}&page={page_num}'
                yield scrapy.Request(url=amazon_search_url, callback=self.parse_search_results, meta={'keyword': keyword, 'page': page_num})

    def parse_product_data(self, response):
        variant_data = re.findall(r'dimensionValuesDisplayData"\s*:\s* ({.+?}),\n', response.text)
        feature_bullets = [bullet.strip() for bullet in response.css("#feature-bullets li ::text").getall()]
        price = response.css('.a-price span[aria-hidden="true"] ::text').get("")
        if not price:
            price = response.css('.a-price .a-offscreen ::text').get("")
        
        for i in range(len(excludeBrands)):
            if excludeBrands[i] == response.css('a.a-link-normal::text').get("").strip():
                yield
            else:
                brand = response.css('a.a-link-normal::text').get("").strip()
        if inOut == 1:
            if response.css('div#a-box#outOfStock ::text').get(""):
                stock = "Out of stock"
            else:
                stock = "In stock"
        else:
            if response.css('div#a-box#outOfStock ::text').get(""):
                yield

        

        yield {
            "name": response.css("#productTitle::text").get("").strip(),
            "price": price,
            "stars": response.css("i[data-hook=average-star-rating] ::text").get("").strip(),
            "rating_count": response.css("div[data-hook=total-review-count] ::text").get("").strip(),
            "brand": brand,
            "feature_bullets": feature_bullets,
            "variant_data": variant_data,
            "stock": stock
        }
    

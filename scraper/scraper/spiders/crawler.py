import scrapy
from scrapy.exceptions import CloseSpider

class CrawlerSpider(scrapy.Spider):
    name = "crawler"
    allowed_domains = ["ecourt.ge"]
    start_urls = ["https://ecourt.ge/Insolvency/InsolvencyDocs?FromDate=&ToDate=&CaseNo=&CouCode=&Partator=&PageNumber=1"]
    page_number = 1

    def parse(self, response):
        if len( response.css('div#InsolvencyDocs ul li.sl') ) == 0:
            raise CloseSpider('მეტი აღარაა')
        
        case = response.css('div#InsolvencyDocs ul li.sl')
        for case in case:
            relative_url = case.css('a ::attr(href)').get()
            case_url = 'https://www.ecourt.ge' + relative_url
            not_to_scrape = ['https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=3015861&instanceId=1', 'https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=2677541&instanceId=1', 'https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=4187203&instanceId=1', 'https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=3452339&instanceId=1', 'https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=4288306&instanceId=1', 'https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=5606897&instanceId=1','https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=6557841&instanceId=1','https://www.ecourt.ge/Insolvency/InsolvencyInner?docId=6328182&instanceId=1']
            if case_url in not_to_scrape:
                yield scrapy.Request(case_url, callback=self.not_parse)
            yield response.follow(case_url, callback = self.parse_case_page)
        self.page_number += 1
        next_page_url = f'https://ecourt.ge/Insolvency/InsolvencyDocs?FromDate=&ToDate=&CaseNo=&CouCode=&Partator=&PageNumber={self.page_number}'
        yield response.follow(next_page_url, callback = self.parse)
    
    def parse_case_page(self, response):
        pdfs = len(response.css('body div div div'))
        yield{
        'მოსამართლე' : response.css('.conversationInnerHeader ul li:nth-of-type(4) span::text').get(),
        'url' : response.url,
        'თარიღი' : response.css('.conversationInnerHeader ul li:nth-of-type(1) span::text').get(),
        'pdf count' : pdfs
        }
    
    
    def not_parse(self, response):
        pass
    
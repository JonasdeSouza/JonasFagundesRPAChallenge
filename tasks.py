from RPA.Browser.Selenium import Selenium
from RPA.Calendar import Calendar
from RPA.Excel.Files import Files
from robocorp.tasks import task
from robocorp import workitems 

from datetime import datetime
import re
import pathlib
import logging
logger = logging.getLogger(__name__)

class Excel(Files):
    def read_excel_worksheet(self, path, worksheet):
        logger.info('Reading excel worksheet')
        try:
            self.open_workbook(path)
        except:
            logger.info("File doesn't exist, creating based on model.")
            self.create_wb_model(path, worksheet)
        finally:
            self.save_workbook()
            self.close_workbook()
            logger.info('Excel file created')

    def create_wb_model(self, path, worksheet):
        self.create_workbook(path=path, fmt="xlsx")
        #self.create_worksheet(worksheet)
        self.rename_worksheet(self.get_active_worksheet(), worksheet)
        self.save_workbook()
        self.open_workbook(path)
        table = {
            "title": ["lorem ipsum"],
            "date":  [],
            "description":  [],
            "picture filename":  [],
            "count search prhases":  [],
            "countain money?":  [],
            }
        self.append_rows_to_worksheet(table, header=True)
        self.auto_size_columns("A", "F")
        self.delete_rows(2)
        self.save_workbook()
        logger.info('Excel file successfully created')

    def append_rows(self, results, path, worksheet):
        logger.info('Appending data to excel file')
        self.open_workbook(path)
        for table in results: 
            self.append_rows_to_worksheet(table, header=False, name=worksheet)
        self.save_workbook()
        self.close_workbook()
        logger.info('Data successfully appended')


class Browser(Selenium):

    tries_counter = 0
    url_size = 0

    topics_list = []
    checkboxes_list = []

    results_list = [[]]
    within_delta = True

    search = None
    sort_by = None
    filter_topics = []
    
    def search(self, search_query: str):
        logger.info('Started searching')
        self.search = search_query
        try:
            search_button = self.find_elements("data:element:search-button")
            self.click_button(search_button)
            search_bar = self.find_element("name:q")
            self.wait_until_element_is_enabled(search_bar)
            #self.click_element(search_bar)
            self.input_text(search_bar, search_query)
            search_button = self.find_elements("data:element:search-submit-button")
            self.click_button(search_button)
        finally:
            self.tries_counter = 0
            logger.info('Finished searching')


    def sort_by(self, sort_type: str):
        logger.info('Started sorting')
        self.sort_by = sort_type
        select = self.find_element("name:s")
        self.wait_until_element_is_enabled(select)
        match sort_type:
            case "Newest":
                wait_for = 1
            case "Oldest":
                wait_for = 2
            case "Relevance":
                wait_for = 0
            case _:
                raise Exception("The sort type '" + sort_type + "' do not exist")
        self.select_from_list_by_label(select, sort_type)
        self.go_to(self.get_location())
        results = self.find_element("class:search-results-module-results-menu")
        self.wait_until_element_is_enabled(results)
        logger.info('Finished sorting')
        
    def filter_topics(self, topics):
        logger.info('Started filtering')
        self.filter_topics = topics
        for i in range(len(topics)):
            self.url_size = len(self.get_location())
            self.update_filters()
            self.checkbox_select(topics[i])
            self.topics_list = None
            self.go_to(self.get_location())
            self.wait_until_element_is_enabled("class:search-results-module-filters-selected")
            logger.info('Filter ' + str(i) + ' applied')
        logger.info('Finished filtering')
        
    def update_filters(self):
        logger.info('Updating filters list')
        try:
            topics_list_element = self.find_element("class:search-filter")
            self.wait_until_element_is_enabled(topics_list_element)
            see_all_button = self.find_element("class:see-all-text")
            self.wait_until_page_contains_element(see_all_button)
            self.wait_until_element_is_enabled(see_all_button)
            self.click_element(see_all_button)
            see_less_span = self.find_element("class:see-less-text")
            self.wait_until_element_is_enabled(see_less_span)
            topics_list_ul = self.find_element("class:search-filter-menu")
            self.topics_list = self.find_elements("//li/div/div/label/span", topics_list_ul)
            self.checkboxes_list = self.find_elements("//li/div/div/label/input", topics_list_ul)
        except:
            logger.warning('Exception occoured while updating filters, trying again...')
            self.update_filters()
            return
        logger.info('Filters list updating complete')

    def checkbox_select(self, topic):
        logger.info('Started checkbox clicking')
        try:
            for topic_index, single_topic in enumerate(self.topics_list):
                if single_topic.text.lower() == topic:
                    checkbox = self.checkboxes_list[topic_index]
                    self.wait_until_element_is_enabled(checkbox)
                    self.click_element(checkbox)
        except:
            logger.warning('Exception occoured while clicking on checkbox, trying again...')
            if self.tries_counter >= 5:
                return
            self.tries_counter += 1
            self.checkbox_select(topic)
            return
        logger.info('Checkbox marked')

    def wait_url_update(self, str, str_start_index):
        logger.info('Waiting url update')
        updated = False
        while updated == False:
            url = self.get_location()
            if str in url[str_start_index-4:]:
                updated == True
                logger.info('Url updated')
                return
        logger.info('Url updated')

    def find_results(self, no_of_months):
        logger.info('Looking for search results')
        while self.within_delta == True:
            self.get_results(no_of_months)
            next_page_button = self.find_element("class:search-results-module-next-page")
            self.wait_until_element_is_enabled(next_page_button)
            if self.find_element("class:search-results-module-page-counts").text[:2] == "10":
                self.within_delta == False
                logger.info('Search result not within specified time delta, stopping')
                break
            self.click_element_when_clickable(next_page_button)
            logger.info('Whole page of results scanned, going to next page')

    def get_results(self, no_of_months):
        logger.info('Getting search results elements')
        results_menu = self.find_element("class:search-results-module-results-menu")
        self.wait_until_element_is_enabled(results_menu)
        results_list = self.find_elements("//li/ps-promo/div")
        for index, result in enumerate(results_list):
            self.within_delta = self.check_date(result, no_of_months)
            if self.within_delta == True:
                self.get_data(result)

    def check_date(self, element, no_of_months):
        logger.info('Checking time delta')
        publication_date = self.find_element("class:promo-timestamp", element)
        publication_date = publication_date.text
        try:
            if "." in publication_date:
                formatted_date = datetime.strptime(publication_date, "%b. %d, %Y")
            else:
                formatted_date = datetime.strptime(publication_date, "%B %d, %Y")
        except:
            formatted_date = datetime.today()

        today = datetime.today()
        months_diff = Calendar().time_difference_in_months(datetime.strftime(formatted_date, "%Y-%m-%d"), datetime.strftime(today, "%Y-%m-%d"))

        if months_diff < no_of_months:
            logger.info('Search result within specified time delta')
            return True
        else:
            return False

    def get_data(self, element):
        logger.info('Formatting data')
        title = self.find_element("class:promo-title", element).text
        formatted_date = self.format_date(element)
        description = self.find_element("class:promo-description", element).text
        picture_filename = self.get_picture(element, title)
        count_search_phrases = self.count_search_phrases(title, description)
        contains_money = self.contains_money(title, description)
        table = {
            "title": [title],
            "date":  [formatted_date],
            "description":  [description],
            "picture filename":  [picture_filename],
            "count search phrases":  [count_search_phrases],
            "countain money?":  [contains_money],
            }
        self.results_list.append(table)
        logger.info('Data formatted')

    def format_date(self, element):
        logger.info('Formatting date')
        publication_date = self.find_element("class:promo-timestamp", element)
        publication_date = publication_date.text
        try:
            if "." in publication_date:
                formatted_date = datetime.strptime(publication_date, "%b. %d, %Y")
            else:
                formatted_date = datetime.strptime(publication_date, "%B %d, %Y")
        except:
            formatted_date = datetime.today()
        finally:
            logger.info('Date formatted')
            return formatted_date
        
    def get_picture(self, element, title):
        logger.info('Getting picture from article')
        picture = self.find_element("class:promo-media", element)
        picture_filename = r"output/" + title.replace(" ", "_") + ".png"
        picture.screenshot(picture_filename)
        logger.info('Picture saved')
        return picture_filename

    def count_search_phrases(self, title, description):
        logger.info('Counting search phrases')
        search_query = self.search.split()
        word_appearances = 0
        for index, word in enumerate(search_query):
            if word in title:
                word_appearances += 1
            if word in description:
                word_appearances += 1
        logger.info('Search phrases counted')
        return word_appearances

    def contains_money(self, title, description):
        logger.info('Searching for money formats')
        money_pattern = r"\$[\d,]+(?:\.\d+)?|\d+\s*(?:dollars|USD)"
        match = re.search(money_pattern, title)
        match = re.search(money_pattern, description)
        if match:
            logger.info('Title or description contains money')
            return True
        else:
            logger.info("Title or description doesn't contain money")
            return False

@task
def setup():
    item = workitems.inputs.current
    payload = dict(item.payload)
    
    arg1 = payload.get("search_query")
    arg2 = payload.get("sort_by")
    arg3 = payload.get("no_of_months")
    arg4 = []
    for item in payload.values():
        if item != arg1 and item != arg2 and item != arg3:
            arg4.append(item)
    print(arg4)
    run_automation(arg1, arg2, arg3, arg4)

def run_automation(search_query:str, sort_by:str, no_of_months:int, topics:list):
    logging.basicConfig(filename=r'output/automation.log', level=logging.INFO)
    logger.info('Started main')
    browser = Browser()
    browser.open_available_browser("https://www.latimes.com/")
    browser.maximize_browser_window()
    browser.search(search_query)
    browser.sort_by(sort_by)
    browser.filter_topics(topics)
    browser.find_results(no_of_months)

    path = str(pathlib.Path(__file__).parent.resolve()) + r"/output/results.xlsx"

    excel_file = Excel()
    excel_file.read_excel_worksheet(path, "data")
    excel_file.append_rows(browser.results_list, path, "data")

    browser.close_browser()
    logger.info('Finished main')

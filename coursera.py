import requests
from bs4 import BeautifulSoup
import openpyxl
import argparse


def fetch_soup_with_content_from_url(url):
    try:
        responce = requests.get(url)
        return BeautifulSoup(responce.content, features="lxml")
    except requests.HTTPError as error:
        print('HTTP Error!')
        print('Response is: {0}'.format(error.response.content))
    except requests.ConnectionError:
        print('Connection failed!')


def get_courses_list(courses_list_size):
    coursera_sitemap_url = 'https://www.coursera.org/sitemap~www~courses.xml'
    soup = fetch_soup_with_content_from_url(coursera_sitemap_url)
    return [tag.text for tag in soup.findAll('loc', limit=courses_list_size)]


def get_courses_info(courses_url_list):
    courses_info = []
    for url in courses_url_list:
        description = get_course_description(url)
        courses_info.append(description)
    return courses_info


def get_course_description(course_url):
    description = {}
    soup = fetch_soup_with_content_from_url(course_url)

    found_tag = soup.find('h1', class_='title display-3-text')
    description['name'] = found_tag.text
    found_tag = soup.find('div', class_='rc-Language')
    description['language'] = found_tag.text
    found_tag = soup.find('div', class_='ratings-text bt3-visible-xs')
    if found_tag:
        description['average_rating'] = found_tag.text
    else:
        description['average_rating'] = 'Unrated'
    found_tag = soup.find('div', class_='startdate rc-StartDateString caption-text')
    words = found_tag.text.split(' ')
    description['start_date'] = words[1] + ' ' + words[2]
    all_found_tags = soup.find_all('div', class_='week')
    description['count_of_week'] = len(all_found_tags)
    return description


def output_courses_info_to_xlsx(filepath, courses_info):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['Name',
                  'Language',
                  'Average rating',
                  'Start date',
                  'Count of week'])

    for course in courses_info:
        sheet.append([course['name'],
                      course['language'],
                      course['average_rating'],
                      course['start_date'],
                      course['count_of_week']])
    workbook.save(filepath)


def parse_and_validate_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--filepath", action="store",
                        help="file for saving to xlsx")
    parser.add_argument("-s", "--listsize", action="store",
                        help="list size for saving")
    args = parser.parse_args()
    filepath = args.filepath

    if not filepath:
        print('File for saving to xlsx not specified!')
        exit()
    try:
        check_file = open(filepath, 'w')
        check_file.close()
    except FileNotFoundError:
        print('Invalid file path!')
        exit()

    if not args.listsize:
        courses_list_size = 20
    else:
        courses_list_size = int(args.listsize)

    return filepath, courses_list_size


if __name__ == '__main__':
    filepath, courses_list_size = parse_and_validate_arguments()

    courses_url_list = get_courses_list(courses_list_size)
    if courses_url_list:
        courses_info = get_courses_info(courses_url_list)
        output_courses_info_to_xlsx(filepath, courses_info)
        print('Courses list wrote successfully into ' + filepath)

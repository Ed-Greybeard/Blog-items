#
# Name:        Skype_Media.py
# Purpose:     Extract details from Skype Messaging cache
# Author:      Chris Ed
# Created:     27/04/2015
# Copyright:   Chris Ed (c) 2015
# Licence:     Open
#

import sqlite3
import logging
import os
from sys import exit
from tkinter.filedialog import askdirectory
from tkinter import Tk

# Little hack to remove tk window
Tk().withdraw()

# Global variables, for it is they
DB_DICT = {}
FOLDER_PATH = ""

logging.basicConfig(level=logging.INFO)


def find_files(source_dir):
    """
    Locates the three databases we require.
    :param source_dir source directory - hopefully a Skype user directory!:
    :return: dict with the following values:
                {
                  cache_db   : cache_db.db path,
                  main_db    : main.db path,
                  storage_db : storage_db.db path
                }
    """
    required_files = ['cache_db.db', 'main.db', "storage_db.db"]
    path_dict = {}
    for root, folders, files in os.walk(source_dir):
        for file_name in files:
            if file_name in required_files:
                if file_name == "cache_db.db":
                    path_dict['cache_db'] = os.path.join(root, file_name)
                elif file_name == "main.db":
                    path_dict['main_db'] = os.path.join(root, file_name)
                elif file_name == "storage_db.db":
                    path_dict['storage_db'] = os.path.join(root, file_name)

    logging.debug(path_dict)
    return path_dict


def get_authors(transfer_dict):
    """
    Extracts authors/senders from the Messages table of main.db
    :param transfer_dict:  returned from get_file_uri_assoc
    :return:  list of useful dialog partners
    """
    file_author = {}
    conn = sqlite3.connect(DB_DICT['main_db'])
    curr = conn.cursor()
    for uri, file_name in transfer_dict.items():
        sql = """
            select author, datetime(timestamp, 'unixepoch'), body_xml
            from Messages
            where body_xml like "%%%s%%"
            order by timestamp
            """ % uri
        logging.debug(sql)
        curr.execute(sql)
        for row in curr:
            file_author[file_name] = [row[1], row[0], get_original_filename(row[2])]
    logging.debug(file_author)
    return file_author


def get_original_filename(body_xml_value):
    """
    Extracts the original filename (v=) from body_xml_value
    :param body_xml_value: returned from body_xml field of Messages table in main.db
    :return: original filename
    """
    logging.debug(body_xml_value)
    # not using the xml module - trying to keep imports down (ha!)
    start_offset = body_xml_value.index("v=\"") + 3
    logging.debug(body_xml_value[start_offset:])
    end_offset = body_xml_value.index("\"/>", start_offset)
    logging.debug(body_xml_value[start_offset:end_offset])
    return body_xml_value[start_offset:end_offset]


def get_sent_uri(file_name):
    """

    :param file_name:   name of the file we're looking for
    :return:            returns URI value for file_name
    """
    conn = sqlite3.connect(DB_DICT['storage_db'])
    curr = conn.cursor()
    file_index = file_name[1:2]
    sql = "select uri from documents where id = " + file_index
    curr.execute(sql)
    uri = ""
    for row in curr:
        uri = row[0]
    logging.debug(uri)
    conn.close()
    return uri


def get_file_uri_assoc():
    """
    Extracts filename and uri associations
    :return: dict containing {uri : filepath} key value pairs
    """
    conn = sqlite3.connect(DB_DICT['cache_db'])
    curr = conn.cursor()
    sql = "select * from assets"
    curr.execute(sql)
    transfer = {}
    for row in curr:
        source = row[1][0]
        index = row[1][1:]
        serialized_data = row[10]
        cache_path = get_cache_file_name(serialized_data)
        if source == "i":
            uri = get_sent_uri(cache_path)
            transfer[uri] = cache_path
        elif source == "u":
            transfer[index] = cache_path
    logging.debug(transfer)
    return transfer


def get_cache_file_name(serialized_data):
    """
    Extracts file name from serialized_data blob
    :param serialized_data: serialized_data field from cache_db
    :return: filename which exists in the blob
    """
    hit_str = b"$CACHE/\\\\"
    logging.debug(len(serialized_data))
    path_offset = (serialized_data.find(hit_str)) + len(hit_str)
    cache_path = serialized_data[path_offset:]
    end_offset = 0
    if b"pimgpsh" in cache_path:
        end_offset = cache_path.index(b"\x01")
    elif b"cimgpsh" in cache_path:
        end_offset = cache_path.index(b"\x08")
    cache_path = str(cache_path[:end_offset], "ascii")
    return cache_path


def generate_html_report(file_auth_dict, acc_name):
    """
    Creates a very basic HTML report
    :param file_auth_dict:  dictionary consist of filename:[author,date_sent] key value pairs
    :return: nuffink
    """
    # TODO Add link to images (maybe requires copying the images locally? Maybe b64 encode?)
    html_report = open("report.html", 'w')
    html_report.write("""<html><body><font face="calibri"><h1>Sample Simple Report</h1>\n""")
    html_report.write("""<h2>Skype Account: %s</h2>\n""" % acc_name)
    html_report.write("""<table border=0 cellpadding=2 cellspacing=2>\n""")
    html_report.write("""<tr><th>Date</th><th>Sender</th><th>Local Filename</th><th>Original Filename</th></tr>\n""")
    for file_name, dateauthor in file_auth_dict.items():
        date_sent = dateauthor[0]
        author = dateauthor[1]
        original_name = dateauthor[2]
        if author == acc_name:
            html_str = """<tr bgcolor=#ffaaaa><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>\n""" % (date_sent,
                                                                                                        author,
                                                                                                        file_name,
                                                                                                        original_name)
        else:
            html_str = """<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>\n""" % (date_sent,
                                                                                      author,
                                                                                      file_name,
                                                                                      original_name)
        html_report.write(html_str)
    html_report.write("""</table></body></html""")
    html_report.close()


def generate_text_report(file_auth_dict, acc_name):
    """
    Creates a tab-delineated text report
    :param file_auth_dict  dictionary consists of filename:[author,date_sent] key value pairs:
    :param acc_name  local Skype account name:
    :return: nuffink
    """
    text_report = open("report.txt", "w")
    text_report.write("""Skype Account: %s\n""" % acc_name)
    text_report.write("""Date\tSender\tFilename\n""")
    for file_name, dateauthor in file_auth_dict.items():
        date_sent = dateauthor[0]
        author = dateauthor[1]
        original_name = dateauthor[2]
        text_str = """%s\t%s\t%s\t%s\n""" % (date_sent, author, file_name, original_name)
        text_report.write(text_str)
    text_report.close()


def get_acc_name():
    """
    Get the local Skype account name
    :return:
    """
    ret_str = ""
    conn = sqlite3.connect(DB_DICT['main_db'])
    curr = conn.cursor()
    curr.execute("select skypename from Accounts")
    for row in curr:
        ret_str = row[0]
    return ret_str


if __name__ == '__main__':
    FOLDER_PATH = askdirectory(title="Locate Skype user directory (AppData\\Roaming\\Skype\\<user>)")
    if FOLDER_PATH == "":
        logging.info("No folder chosen, script cancelled")
        exit()
    logging.info("""Input folder: %s""" % FOLDER_PATH)
    DB_DICT = find_files(FOLDER_PATH)
    logging.debug(len(DB_DICT))
    logging.debug("attempting to search for received files")
    transfer_dict = get_file_uri_assoc()
    logging.debug(len(transfer_dict))
    logging.debug("Find Received authors")
    file_author_dict = get_authors(transfer_dict)
    logging.debug(len(file_author_dict))
    # Will do both reports because we can
    generate_html_report(file_author_dict, get_acc_name())
    generate_text_report(file_author_dict, get_acc_name())

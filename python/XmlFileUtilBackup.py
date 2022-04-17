__author__ = 'Alex Bulavin'
# Backup file for XmlFileUtil.py

#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
from Log import Log
import xml.dom.minidom
import re


class XmlFileUtil:
    """android strings.xml file util"""

    @staticmethod
    def writeToFile(keys, values, directory, filename, additional):
        if not os.path.exists(directory):
            os.makedirs(directory)

        fo = open(directory + "/" + filename, "wb")

        string_encoding = '<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<resources>\n'

        string_encoding_bytes = bytes(string_encoding, encoding="utf8")
        fo.write(string_encoding_bytes)

        for x in range(len(keys)):
            if values[x] is None or values[x] == '':
                Log.error("Key:" + keys[x] +
                          "\'s value is None. Index:" + str(x + 1))
                continue

            key = keys[x].strip()
            value = re.sub(r'(%\d\$)(@)', r'\1s', values[x])
            content = "   <string name=\"" + key + "\">" + value + "</string>\n"
            content_bytes = bytes(content, encoding='utf8')
            fo.write(content_bytes)

        if additional is not None:
            fo.write(additional)

        fo.write(bytes("</resources>", encoding='utf8'))
        fo.close()

    @staticmethod
    def getKeysAndValues(path):  # path here should include full path to the xml file that need to be parse
        keys = []
        values = []
        if path is None:
            Log.error('file path is None')
            return

        doc = xml.dom.minidom.parse(path)
        root = doc.documentElement
        itemlist = root.getElementsByTagName('string')

        for index in range(len(itemlist)):
            item = itemlist[index]
            key = item.getAttribute("name")
            value = item.firstChild.nodeValue
            keys.append(key)
            values.append(value)

        return keys, values

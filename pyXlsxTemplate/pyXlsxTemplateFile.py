import os
import zipfile
import xml.dom.minidom

class XlsxFile(object):
    """
    An Excel 2007 xlsx file that can be modified without losing formatting, formulas, and charts.

    :param filename: The name of the xlsx file.
    """
    def __init__(self, filename=None):
        self.files = {}
        self.worksheets = {}
        self.strings = None
        if filename:
            self.load(filename)

    def load(self, filename):
        """
        Loads an xlsx file.

        :param filename: The name of the xlsx file.
        """
        xlsxfile = zipfile.ZipFile(filename, 'r')
        self.files = { name: xlsxfile.read(name) for name in xlsxfile.namelist() }
        self.worksheets = {
            os.path.basename(name): XlsxWorksheet(self, name)
            for name in self.files
            if os.path.dirname(name.lower()) == 'xl/worksheets'
        }
        self.strings = XlsxSharedStrings(self, 'xl/sharedStrings.xml')
        xlsxfile.close()

    def save(self, filename):
        """
        Saves an xlsx file.

        :param filename: The name of the file to save to.
        """
        xlsxfile = zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED)
        self.strings.save()
        for worksheet in self.worksheets.values():
            worksheet.save()
        for name in self.files:
            xlsxfile.writestr(name, self.files[name])
        xlsxfile.close()

    def loadxml(self, name):
        """
        Loads an XML document from the xlsx file archive.

        :param name: The name of the archived file (full path in the compressed archive).
        :returns: An :class:`xml.dom.minidom` object.
        """
        return xml.dom.minidom.parseString(self.files[name])

    def savexml(self, name, xmldoc):
        """
        Updates the the file in the xlsx file archive with the :class:`xml.dom.minicom` XML object.

        :param name: The name of the archived file (full path in the compressed archive).
        :param xmldoc: The :class:`xml.dom.minitom` object to be saved.
        """
        self.files[name] = xmldoc.toxml()

    def resetAllFormulas(self):
        """
        Removes the values from all of the formulas in the xlsx file so they will be recalculated when the file is opened again.
        """
        for worksheet in self.worksheets.values():
            worksheet.resetFormulas()


class XlsxXml(object):
    def __init__(self, template, name):
        self.template = template
        self.load(name)

    def load(self, name):
        self.name = name
        self.xmldoc = self.template.loadxml(name)

    def save(self):
        self.template.savexml(self.name, self.xmldoc)


class XlsxSharedStrings(XlsxXml):
    def __init__(self, template, name):
        super(XlsxSharedStrings, self).__init__(template, name)

    def load(self, name):
        super(XlsxSharedStrings, self).load(name)
        self.strings = self.xmldoc.getElementsByTagName('t')

    def getString(self, index):
        return self.strings[index].childNodes[0].data

    def setString(self, index, value):
        self.strings[index].childNodes[0].data = unicode(value)


class XlsxWorksheet(XlsxXml):
    def __init__(self, template, name):
        super(XlsxWorksheet, self).__init__(template, name)

    def load(self, name):
        super(XlsxWorksheet, self).load(name)
        self.cells = {}
        for cell_element in self.xmldoc.getElementsByTagName('c'):
            if cell_element.getAttribute('t') == 's':
                cell = XlsxStringCell(self, cell_element)
            else:
                cell = XlsxCell(self, cell_element)
            self.cells[cell_element.getAttribute('r')] = cell

    def resetFormulas(self):
        for formula_element in self.xmldoc.getElementsByTagName('f'):
            cell_element = formula_element.parentNode
            for value_element in cell_element.getElementsByTagName('v'):
                cell_element.removeChild(value_element)


class XlsxCell(object):
    def __init__(self, worksheet, el):
        self.worksheet = worksheet
        self.el = el

    def getValueElement(self):
        try:
            return self.el.getElementsByTagName('v')[0]
        except IndexError:
            return None

    @property
    def template(self):
        return self.worksheet.template

    @property
    def value(self):
        """
        The contents of the cell.
        """
        value_element = self.getValueElement()
        if value_element:
            return value_element.childNodes[0].data
        else:
            return None

    @value.setter
    def value(self, value):
        value_element = self.getValueElement()
        if value_element:
            value_element.childNodes[0].data = unicode(value)
        else:
            value_element = self.worksheet.xmldoc.createElement('v')
            value_element.appendChild(
                self.worksheet.xmldoc.createTextNode(unicode(value))
            )
            self.el.appendChild(value_element)


class XlsxStringCell(XlsxCell):
    def __init__(self, worksheet, el):
        super(XlsxStringCell, self).__init__(worksheet, el)

    @property
    def value(self):
        """
        The string contents of the cell.
        """
        string_index = int(super(XlsxStringCell, self).value)
        return self.template.strings.getString(string_index)

    @value.setter
    def value(self, value):
        string_index = int(super(XlsxStringCell, self).value)
        self.template.strings.setString(string_index, unicode(value))

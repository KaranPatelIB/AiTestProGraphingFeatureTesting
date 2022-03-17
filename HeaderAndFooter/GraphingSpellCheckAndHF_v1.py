# =====================================================================================================
#
# ██ ██████  ███████  █████  ██████  ██    ██ ████████ ███████ ███████      ██ ███    ██  ██████
# ██ ██   ██ ██      ██   ██ ██   ██  ██  ██     ██    ██      ██           ██ ████   ██ ██
# ██ ██   ██ █████   ███████ ██████    ████      ██    █████   ███████      ██ ██ ██  ██ ██
# ██ ██   ██ ██      ██   ██ ██   ██    ██       ██    ██           ██      ██ ██  ██ ██ ██
# ██ ██████  ███████ ██   ██ ██████     ██       ██    ███████ ███████      ██ ██   ████  ██████ ██
#
#
# Project: AiTestPro - Spell Check Integration
# Description: Added spell check functionality and updated Headers and Footer
# By: Karan Patel
# Version History: 1.0 2022-03-10
# Notes: No need for graohviz invisable nodes (thi version will add margins)
# =====================================================================================================


import graphviz
import fitz
import textwrap
import shutil
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from PyPDF2.pdf import PageObject
from PIL import Image, ImageFont, ImageDraw
from fpdf import FPDF
from spellchecker import SpellChecker
import re
from collections import Counter
import os
import time


# ===================== Just need to add this section of the code to bottom of script ==================================

# Notes: No Changes need to be made to this. It is used to create tables,
# It can also be put in a separate file with the class imported in the main script

class PDF(FPDF):
    def create_table(self, table_data, title='', data_size=10, title_size=12, align_data='L', align_header='L',
                     cell_width='even', x_start='x_default', emphasize_data=[], emphasize_style=None,
                     emphasize_color=(0, 0, 0)):
        """
        table_data:
                    list of lists with first element being list of headers
        title:
                    (Optional) title of table (optional)
        data_size:
                    the font size of table data
        title_size:
                    the font size fo the title of the table
        align_data:
                    align table data
                    L = left align
                    C = center align
                    R = right align
        align_header:
                    align table data
                    L = left align
                    C = center align
                    R = right align
        cell_width:
                    even: evenly distribute cell/column width
                    uneven: base cell size on lenght of cell/column items
                    int: int value for width of each cell/column
                    list of ints: list equal to number of columns with the widht of each cell / column
        x_start:
                    where the left edge of table should start
        emphasize_data:
                    which data elements are to be emphasized - pass as list
                    emphasize_style: the font style you want emphaized data to take
                    emphasize_color: emphasize color (if other than black)

        """
        default_style = self.font_style
        if emphasize_style == None:
            emphasize_style = default_style

        # default_font = self.font_family
        # default_size = self.font_size_pt
        # default_style = self.font_style
        # default_color = self.color # This does not work

        # Get Width of Columns
        def get_col_widths():
            col_width = cell_width
            if col_width == 'even':
                col_width = self.epw / len(data[
                                               0]) - 1  # distribute content evenly   # epw = effective page width (width of page not including margins)
            elif col_width == 'uneven':
                col_widths = []

                # searching through columns for largest sized cell (not rows but cols)
                for col in range(len(table_data[0])):  # for every row
                    longest = 0
                    for row in range(len(table_data)):
                        cell_value = str(table_data[row][col])
                        value_length = self.get_string_width(cell_value)
                        if value_length > longest:
                            longest = value_length
                    col_widths.append(longest + 4)  # add 4 for padding
                col_width = col_widths
                ### compare columns

            elif isinstance(cell_width, list):
                col_width = cell_width  # TODO: convert all items in list to int
            else:
                # TODO: Add try catch
                col_width = int(col_width)
            return col_width

        # Convert dict to lol
        # Why? because i built it with lol first and added dict func after
        # Is there performance differences?
        if isinstance(table_data, dict):
            header = [key for key in table_data]
            data = []
            for key in table_data:
                value = table_data[key]
                data.append(value)
            # need to zip so data is in correct format (first, second, third --> not first, first, first)
            data = [list(a) for a in zip(*data)]

        else:
            header = table_data[0]
            data = table_data[1:]

        line_height = self.font_size * 2.5

        col_width = get_col_widths()
        self.set_font(size=title_size + 10)
        self.set_font(style='B')
        self.set_text_color(2, 79, 151)

        # Get starting position of x
        # Determin width of table to get x starting point for centred table
        if x_start == 'C':
            table_width = 0
            if isinstance(col_width, list):
                for width in col_width:
                    table_width += width
            else:  # need to multiply cell width by number of cells to get table width
                table_width = col_width * len(table_data[0])
            # Get x start by subtracting table width from pdf width and divide by 2 (margins)
            margin_width = self.w - table_width
            # TODO: Check if table_width is larger than pdf width

            center_table = margin_width / 2  # only want width of left margin not both
            x_start = center_table
            self.set_x(x_start)
        elif isinstance(x_start, int):
            self.set_x(x_start)
        elif x_start == 'x_default':
            x_start = self.set_x(self.l_margin)

        # TABLE CREATION #

        # add title
        if title != '':
            self.multi_cell(0, line_height, title, border=0, align='C', ln=3, max_line_height=self.font_size,
                            markdown=True)
            self.ln(line_height)  # move cursor back to the left margin

        self.set_font(size=data_size)
        self.set_font(style='B')
        self.set_text_color(229, 125, 28)
        # self.set_fill_color(255,255,255)
        # add header
        y1 = self.get_y()
        if x_start:
            x_left = x_start
        else:
            x_left = self.get_x()
        x_right = self.epw + x_left
        if not isinstance(col_width, list):
            if x_start:
                self.set_x(x_start)
            for datum in header:
                self.multi_cell(col_width, line_height, datum, border=0, align=align_header, ln=3,
                                max_line_height=self.font_size)
                x_right = self.get_x()
            self.ln(line_height)  # move cursor back to the left margin
            y2 = self.get_y()
            self.line(x_left, y1, x_right, y1)
            self.line(x_left, y2, x_right, y2)

            for row in data:
                if x_start:  # not sure if I need this
                    self.set_x(x_start)
                for datum in row:
                    if datum in emphasize_data:
                        self.set_text_color(*emphasize_color)
                        self.set_font(style=emphasize_style)
                        self.multi_cell(col_width, line_height, datum, border=0, align=align_data, ln=3,
                                        max_line_height=self.font_size)
                        self.set_text_color(0, 0, 0)
                        self.set_font(style=default_style)
                    else:
                        self.multi_cell(col_width, line_height, datum, border=0, align=align_data, ln=3,
                                        max_line_height=self.font_size)  # ln = 3 - move cursor to right with same vertical offset # this uses an object named self
                self.ln(line_height)  # move cursor back to the left margin

        else:
            if x_start:
                self.set_x(x_start)
            for i in range(len(header)):
                datum = header[i]
                self.multi_cell(col_width[i], line_height, datum, border=0, align=align_header, ln=3,
                                max_line_height=self.font_size)
                x_right = self.get_x()
            self.ln(line_height)  # move cursor back to the left margin
            y2 = self.get_y()
            self.line(x_left, y1, x_right, y1)
            self.line(x_left, y2, x_right, y2)

            self.set_font()
            self.set_text_color(0, 0, 0)

            for i in range(len(data)):
                if x_start:
                    self.set_x(x_start)
                row = data[i]
                for i in range(len(row)):
                    datum = row[i]
                    if not isinstance(datum, str):
                        datum = str(datum)
                    adjusted_col_width = col_width[i]
                    if datum in emphasize_data:
                        self.set_text_color(*emphasize_color)
                        self.set_font(style=emphasize_style)
                        self.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, ln=3,
                                        max_line_height=self.font_size)
                        self.set_text_color(0, 0, 0)
                        self.set_font(style=default_style)
                    else:
                        self.multi_cell(adjusted_col_width, line_height, datum, border=0, align=align_data, ln=3,
                                        max_line_height=self.font_size)  # ln = 3 - move cursor to right with same vertical offset # this uses an object named self
                self.ln(line_height)  # move cursor back to the left margin
        y3 = self.get_y()
        # self.line(x_left, y3, x_right, y3)


# ===================== Just need to Add Code to Bottom of script  ==================================

# Note: Spell Check Class

# testWord1 = 'Conceed Testing play adfasdfewe'
# testWord2 = 'Hi mi namee is Testingg karan this is testt'

WORDS = Counter(re.findall(r'\w+', open('big.txt').read().lower()))


class norvigSpellChecker:

    @staticmethod
    def P(word, N=sum(WORDS.values())):
        # "Probability of `word`."
        return WORDS[word] / N

    @staticmethod
    def correction(word):
        # "Most probable spelling correction for word."
        likely = max(norvigSpellChecker.candidates(word), key=norvigSpellChecker.P)
        if likely == word:
            return "*unknown*"
        return likely

    @staticmethod
    def candidates(word):
        # "Generate possible spelling corrections for word."
        return norvigSpellChecker.known([word]) or norvigSpellChecker.known(
            norvigSpellChecker.edits1(word)) or norvigSpellChecker.known(norvigSpellChecker.edits2(word)) or [word]

    @staticmethod
    def known(words):
        # "The subset of `words` that appear in the dictionary of WORDS."
        return set(w for w in words if w in WORDS)

    @staticmethod
    def edits1(word):
        # "All edits that are one edit away from `word`."
        letters = 'abcdefghijklmnopqrstuvwxyz'
        splits = [(word[:i], word[i:]) for i in range(len(word) + 1)]
        deletes = [L + R[1:] for L, R in splits if R]
        transposes = [L + R[1] + R[0] + R[2:] for L, R in splits if len(R) > 1]
        replaces = [L + c + R[1:] for L, R in splits if R for c in letters]
        inserts = [L + c + R for L, R in splits for c in letters]
        return set(deletes + transposes + replaces + inserts)

    @staticmethod
    def edits2(word):
        # "All edits that are two edits away from `word`."
        return (e2 for e1 in norvigSpellChecker.edits1(word) for e2 in norvigSpellChecker.edits1(e1))


class Spell:

    def __init__(self, fileLocationPath):
        self.pathToFile = fileLocationPath
        self.tableRows = [["MISSPELLED", "SUGGESTIONS", "SHEETNAME", "ROW:COLUMN", "INDEX", "XPATH"]]

    def spell(self, string, sheetName, cellValueRow, cellValueColumn, sheetIndex, xPath):
        # spellChecker = SpellChecker()
        # missSpelled = spellChecker.unknown(spellChecker.split_words(string))
        # missSpelled = set(string.split())
        missSpelled = spellchecker.split_words

        for missSpelledWord in missSpelled:
            # Get the one `most likely` answer
            corrections = [norvigSpellChecker.correction(missSpelledWord)]
            if corrections == ['*unknown*']:
                entry = "UNKNOWN WORD"
            else:
                entry = str(corrections)
            # Get a list of `likely` options
            # print((spell.candidates(missSpelledWord)).)

            temp = [missSpelledWord, entry, sheetName, str(str(cellValueRow) + ":" + str(cellValueColumn)),
                    sheetIndex, xPath]
            self.tableRows.append(temp)
            # print(temp)

    def createPDF(self):
        pdf = PDF()
        pdf.set_top_margin(20)
        pdf.add_page(orientation='L')
        pdf.set_font("Arial", size=9)

        pdf.create_table(table_data=self.tableRows, title="SPELL CHECK", cell_width=[30, 30, 30, 25, 20, 150],
                         align_data='L',
                         align_header='L', x_start='C', data_size=8, title_size=10)

        pdf.ln()
        pdf.output(self.pathToFile)  # os.path.join(self.pathToFile, 'table_class.pdf'))


# '''
start_time = time.time()
# Pass resulting pdf location
sheet = Spell('C:\\Users\\email\\ideabytes\\HeaderAndFooter\\spellCheck.pdf')

# String | Sheet Name | Sheet Row | Sheet Column | Sheet Index | Xpath          (You can call as many times as you want)
sheet.spell("If Yuo’re Albe To Raed Tihs, You Might Have Typoglycemia", "Testing1", 1, 2, 20,
            "/dicsa[][][][]sdsfdf[a]df[]asd[f][sda]f[xc]")
sheet.spell(
    "Aoccdrnig to a rscheearch at Cmabrigde Uinervtisy, it deosn’t mttaer in waht oredr the ltteers in a wrod are, the olny iprmoetnt tihng is taht the frist and lsat ltteer be at the rghit pclae. The rset can be a toatl mses and you can sitll raed it wouthit porbelm. Tihs is bcuseae the huamn mnid deos not raed ervey lteter by istlef, but the wrod as a wlohe.",
    "Testing2", 1, 2, 20, "/dicsasd[][][][sfdf[a]df[]asd[f][sda]f[xc]")

for i in range(100):
    sheet.spell("Testting ABCDEFADS Conceed", "Testing3", 1, 2, 20, "/dicsasdsfdf[a]df[][][][[]asd[f][sda]f[xc]")
# When you are done simpy call this method to create pdf at previously given location
sheet.createPDF()

end_time = time.time()
print("Total Spell Check Time: " + str(end_time - start_time))
# '''

# ===================== Update Header and Footer Class ==================================

# Note : Header and Footer class has changed internally, so you will need to replace old class with new class.
# As per arguments you will just need to pass spell check pdf location as last argument (check below)


# '''
def make_graph():
    graph = graphviz.Graph(name="ex")

    # invisible nodes to add margin
    graph.node('header', "", shape='none')
    graph.node('footer', "", shape='none')

    graph.node('head', "http://testaws.dgsms.ca/ OpenUrl", shape='note')
    graph.node('A1', 'Forgot Password', shape='cds', color='chartreuse')
    graph.node('A2', 'Trial Registration', shape='cds', color='chartreuse')

    with graph.subgraph(name="Cluster_B1") as subGraph:
        # subGraph.attr()
        subGraph.node('B1', 'username', shape='underline')
        subGraph.node('B2', 'userpass', shape='underline')
        subGraph.node('B3', 'submit Login', shape='cds', color='chartreuse')
        subGraph.edges([('B1', 'B2'), ('B2', 'B3')])

    graph.node('C1', 'http://testaws.dgsms.ca/LoginAction?action=login', shape='note')

    graph.edges([('head', 'A1'), ('head', 'A2'), ('head', 'B1'), ('B3', 'C1')])

    graph.render('graphExampleComplete', view=True)


# '''


class header_and_footer:

    def __init__(self, input_pdf, overwrite_pdf, spellCheck_pdf):
        self.input_pdf = input_pdf
        self.overwrite_pdf = overwrite_pdf
        self.spellCheck_pdf = spellCheck_pdf

    def create_header(self, page_width, page_height):
        scale = 3
        header = Image.new(mode="RGB", size=(round(page_width * scale), round(75 * scale)),
                           color=(255, 255, 255))  # color=(37, 166, 218))
        # color=(37, 166, 218)) - color=(255,255,255))#
        logo = Image.open(self.overwrite_pdf.get("full_logo_path"))
        logo.thumbnail((100 * scale, 100 * scale))

        # Left Side
        left_side_text = ("Sheet Name: \n> " + split_string(self.overwrite_pdf.get("sheet_info")[0]) +
                          "Pre Requisites: \n> " + split_string(self.overwrite_pdf.get("pre_requisites")[0]) +
                          split_string("Step Number: " + self.overwrite_pdf.get("step_number")[0]))

        # Center
        header.paste(logo, (round(header.width / 2 - logo.width / 2), round(header.height / 2 - logo.height / 2)))

        # Right Side
        right_side_text = (split_string("Version Number: " + self.overwrite_pdf.get("version_number")) +
                           split_string("Server: " + self.overwrite_pdf.get("server")))

        # Font Setup
        font_size = 8 * scale
        font = ImageFont.truetype("arial.ttf", font_size)
        # rightW, rightH = font.getsize(right_side_text)
        leftH = header.height
        leftW = header.width

        while leftH > (header.height - 5) / scale:
            font_size -= 1
            font = ImageFont.truetype("arial.ttf", font_size)
            leftW, leftH = font.getsize(left_side_text)

        d = ImageDraw.Draw(header)

        d.text((10 * scale, round((header.height / 2) - (leftH * scale))), left_side_text, font=font, fill=(0, 0, 0))
        d.text((header.width - 38 * 2.5 * scale, round((header.height / 2) - (leftH * scale))), right_side_text,
               font=font, fill=(0, 0, 0))

        header.save("header.png")
        return header.width / scale, header.height / scale

    def create_footer(self, page_width, page_height):
        scale = 3
        footer = Image.new(mode="RGB", size=(round(page_width * scale), round(75 * scale)),
                           color=(255, 255, 255))  # color=(37, 166, 218))
        # color=(37, 166, 218)) - color=(255,255,255))#
        logo = Image.open(self.overwrite_pdf.get("small_logo_path"))
        logo.thumbnail((50 * scale, 50 * scale))

        # Left Side
        left_side_text = ("Created On: " + self.overwrite_pdf.get("created_on") +
                          ("\nPage Number: " + self.overwrite_pdf.get("page_number")))

        # Right Side
        footer.paste(logo, (round(footer.width - logo.width - 10 * scale), round(footer.height / 2 - logo.height / 2)))

        # Font Setup
        font_size = 10
        font = ImageFont.truetype("arial.ttf", font_size * scale)
        leftH = footer.height
        leftW = footer.width

        while leftH > (footer.height - 5) / scale:
            font_size -= 1
            font = ImageFont.truetype("arial.ttf", font_size * scale)
            leftW, leftH = font.getsize(left_side_text)

        d = ImageDraw.Draw(footer)

        d.text((10 * scale, round((footer.height / 2) - (5 * scale))), left_side_text, font=font, fill=(0, 0, 0))

        footer.save("footer.png")
        return footer.width / scale, footer.height / scale

    def create_spell_header(self, page_width, page_height):
        scale = 3
        header = Image.new(mode="RGB", size=(round(page_width * scale), round(50 * scale)),
                           color=(255, 255, 255))  # color=(37, 166, 218))
        # color=(37, 166, 218)) - color=(255,255,255))#
        logo = Image.open(self.overwrite_pdf.get("full_logo_path"))
        logo.thumbnail((100 * scale, 100 * scale))

        # Center
        header.paste(logo, (40, round(header.height / 2 - logo.height / 2) + 5))

        # Right Side
        right_side_text = ("Version Number: " + split_string(self.overwrite_pdf.get("version_number")))

        # Font Setup
        font_size = 8 * scale
        font = ImageFont.truetype("arial.ttf", font_size)
        # rightW, rightH = font.getsize(right_side_text)
        leftH = header.height
        leftW = header.width

        while leftH > (header.height - 5) / scale:
            font_size -= 1
            font = ImageFont.truetype("arial.ttf", font_size)
            leftW, leftH = font.getsize(right_side_text)

        d = ImageDraw.Draw(header)

        d.text((header.width - 38 * 2.5 * scale, round((header.height / 2) - (leftH * scale))), right_side_text,
               font=font, fill=(0, 0, 0))

        header.save("header.png")
        return header.width / scale, header.height / scale

    def create_spell_footer(self, page_width, page_height, page_counter):
        scale = 3
        footer = Image.new(mode="RGB", size=(round(page_width * scale), round(50 * scale)),
                           color=(255, 255, 255))  # color=(37, 166, 218))
        # color=(37, 166, 218)) - color=(255,255,255))#

        logo = Image.open(self.overwrite_pdf.get("small_logo_path"))
        logo.thumbnail((40 * scale, 40 * scale))

        # Left Side
        left_side_text = ("Created On: " + self.overwrite_pdf.get("created_on") +
                          ("\nPage Number: " + str(page_counter)))

        # Right Side
        footer.paste(logo, (
        round(footer.width - logo.width - 10 * scale - 20), round(footer.height / 2 - logo.height / 2 + 2)))

        # Font Setup
        font_size = 10
        font = ImageFont.truetype("arial.ttf", font_size * scale)
        leftH = footer.height
        leftW = footer.width

        while leftH > (footer.height - 5) / scale:
            font_size -= 1
            font = ImageFont.truetype("arial.ttf", font_size * scale)
            leftW, leftH = font.getsize(left_side_text)

        w, h = 220, 190
        shape = [(40, 0), (2470, 0)]

        d = ImageDraw.Draw(footer)
        d.text((10 * scale + 10, round((footer.height / 2) - (5 * scale))), left_side_text, font=font, fill=(0, 0, 0))
        d.line(shape, width=3, fill="#000000")

        footer.save("footer.png")
        return footer.width / scale, footer.height / scale

    def draw_header_and_footer(self):

        # ----------- Add Margin By Resize ------------

        start_time = time.time()

        docLocation = open(self.input_pdf, 'rb')
        doc = PdfFileReader(docLocation)

        newDoc = PdfFileWriter()
        for pageCounter in range(0, doc.getNumPages()):
            originalPage = doc.getPage(0)
            newPage = PageObject.createBlankPage(None, originalPage.mediaBox.getWidth(),
                                                 originalPage.mediaBox.getHeight() + 100)
            newPage.mergeScaledTranslatedPage(originalPage, 1, 0, 50)
            newDoc.addPage(newPage)
        output = open('Resize.pdf', 'wb')
        newDoc.write(output)
        docLocation.close()
        output.close()
        os.remove(self.input_pdf)
        os.rename('Resize.pdf', self.input_pdf)
        # shutil.copyfile('Resize.pdf', self.input_pdf)
        # os.remove('Resize.pdf')
        end_time = time.time()

        print("Total Margin Resize Time: " + str(end_time - start_time))

        # ----------- Add Header And Footer ------------

        full_logo_file = self.overwrite_pdf.get("full_logo_path")
        small_logo_file = self.overwrite_pdf.get("small_logo_path")
        doc = fitz.open(self.input_pdf)

        page_counter = 0
        for page in doc:
            page_width = round(doc[page_counter].rect.width)
            page_height = round(doc[page_counter].rect.height)

            headerW, headerH = header_and_footer.create_header(self, page_width, page_height)
            footerW, footerH = header_and_footer.create_footer(self, page_width, page_height)
            header_rect = fitz.Rect(0, 0, page_width, headerH + 10)
            footer_rect = fitz.Rect(0, page_height - footerH - 10, page_width, page_height)
            page.insert_image(header_rect, filename="header.png")
            page.insert_image(footer_rect, filename="footer.png")
            page_counter += 1
            self.overwrite_pdf["page_number"] = str(page_counter + 1)
            self.overwrite_pdf["sheet_info"].pop(0)
            self.overwrite_pdf["pre_requisites"].pop(0)
            self.overwrite_pdf["step_number"].pop(0)

        # ----------- Add Spell Check ------------

        totalPages = page_counter + 1
        page_counter = 0
        doc.saveIncr()
        doc.close()
        spellDoc = fitz.open(self.spellCheck_pdf)
        for page in spellDoc:
            page_width = round(spellDoc[page_counter].rect.width)
            page_height = round(spellDoc[page_counter].rect.height)

            headerW, headerH = header_and_footer.create_spell_header(self, page_width, page_height)
            footerW, footerH = header_and_footer.create_spell_footer(self, page_width, page_height, totalPages)
            header_rect = fitz.Rect(0, 0, page_width, headerH + 10)
            footer_rect = fitz.Rect(0, page_height - footerH - 10, page_width, page_height)
            page.insert_image(header_rect, filename="header.png")
            page.insert_image(footer_rect, filename="footer.png")
            page_counter += 1
            totalPages += 1

        spellDoc.saveIncr()
        spellDoc.close()
        os.remove("header.png")
        os.remove("footer.png")

        shutil.copyfile(self.input_pdf, "temp.pdf")
        os.remove(self.input_pdf)

        merger = PdfFileMerger()

        merger.append("temp.pdf")
        merger.append(self.spellCheck_pdf)

        merger.write(self.input_pdf)
        merger.close()
        os.remove("temp.pdf")
        os.remove(self.spellCheck_pdf)


def split_string(text):
    n = 32
    finalWrap = ''
    for x in range(len(textwrap.wrap(text, n, break_long_words=False))):
        finalWrap += textwrap.wrap(text, n, break_long_words=False)[x] + '\n'
    return finalWrap


# ---------------- What you need to use -----------------


def add_header_and_footer(input_file, sheet_info, pre_requisites, step_number, full_logo_path, small_logo_path,
                          version_number,
                          server, created_on, spellCheck_pdf):
    info_pdf = {
        "sheet_info": sheet_info,
        "pre_requisites": pre_requisites,
        "step_number": step_number,
        "full_logo_path": full_logo_path,
        "small_logo_path": small_logo_path,
        "version_number": version_number,
        "server": server,
        "created_on": created_on,
        "page_number": "1",
    }

    operation = header_and_footer(input_file, info_pdf, spellCheck_pdf)
    operation.draw_header_and_footer()

make_graph()
#
# # **Changes**
# # Add Spell Check Pdf as Last Argument
add_header_and_footer(
    "graphExampleComplete.pdf",
    ["Testing Sheet Page 1", "Testing Sheet Page 2", "Testing Sheet Page 3", "Testing Sheet Page 4",
     "Testing Sheet Page 5", "Testing Sheet Page 6"],
    ["pre_requisites Sheet Page 1", "pre_requisites Sheet Page 2", "pre_requisites Sheet Page 3",
     "pre_requisites Sheet Page 4", "pre_requisites Sheet Page 5", "pre_requisites Sheet Page 6"],
    ["102", "202", "302", "402", "502", "602"],
    "aiProTextLogo.jpg",
    "aiProLogo.jpg",
    "v1.1",
    "Canada East",
    "2022-02-17 13:00",
    "spellCheck.pdf")

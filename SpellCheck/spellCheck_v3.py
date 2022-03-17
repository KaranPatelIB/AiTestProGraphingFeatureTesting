from spellchecker import SpellChecker
from create_table_fpdf2 import PDF

testWord1 = 'Conceed Testing play adfasdfewe'
testWord2 = 'Hi mi namee is Testingg karan this is testt'

spell = SpellChecker()

# words = spell.split_words("this sentnce has misspelled werds")

# misspelled = spell.unknown(words)

# for word in misspelled:
#     # Get the one `most likely` answer
#     print(spell.correction(word))
#
#     # Get a list of `likely` options
#     print(spell.candidates(word))


class Spell:

    def __init__(self, fileLocationPath):
        self.pathToFile = fileLocationPath
        self.tableRows = [["MISSPELLED", "SUGGESTIONS", "SHEETNAME", "ROW:COLUMN", "INDEX", "XPATH"]]

    def spell(self, string, sheetName, cellValueRow, cellValueColumn, sheetIndex, xPath):
        spellChecker = SpellChecker()
        missSpelled = spellChecker.unknown(spellChecker.split_words(string))

        for missSpelledWord in missSpelled:
            # Get the one `most likely` answer
            corrections = [spell.correction(missSpelledWord)]
            # Get a list of `likely` options
            # print((spell.candidates(missSpelledWord)).)

            temp = [missSpelledWord, str(corrections), sheetName, str(str(cellValueRow) + ":" + str(cellValueColumn)),
                    sheetIndex, xPath]
            self.tableRows.append(temp)
            # print(temp)

    def createPDF(self):
        pdf = PDF()
        pdf.set_top_margin(20)
        pdf.add_page(orientation='L')
        pdf.set_font("Arial", size=9)

        pdf.create_table(table_data=self.tableRows, title="SPELL CHECK", cell_width=[30,30,30,22,20,150], align_data='L',
                         align_header='C', x_start='C', data_size=8, title_size=10)

        pdf.ln()
        pdf.output(self.pathToFile) #os.path.join(self.pathToFile, 'table_class.pdf'))


# Pass resulting pdf location
sheet = Spell('C:\\Users\\email\\ideabytes\\HeaderAndFooter\\spellCheck.pdf')

# String | Sheet Name | Sheet Row | Sheet Column | Sheet Index | Xpath
sheet.spell("5G korrectud", "Testing1", 1, 2, 20, "/dicsa[][][][]sdsfdf[a]df[]asd[f][sda]f[xc]/dicsa[][][][]sdsfdf[a]df[]asd[f][sda]f[xc]/dicsa[][][][]sdsfdf[a]df[]asd[f][sda]f[xc]/dicsa[][][][]sdsfdf[a]df[]asd[f][sda]f[xc]")
sheet.spell("hi to a rscheearch at Cmabrigde Uinervtisy, it deosn’t mttaer in waht oredr the ltteers in a wrod are, the olny iprmoetnt tihng is taht the frist and lsat ltteer be at the rghit pclae. The rset can be a toatl mses and you can sitll raed it wouthit porbelm. Tihs is bcuseae the huamn mnid deos not raed ervey lteter by istlef, but the wrod as a wlohe.", "Testing2", 1, 2, 20, "/dicsasd[][][][sfdf[a]df[]asd[f][sda]f[xc]")
sheet.spell("Testting ABCDEFADS Conceed", "Testing3", 1, 2, 20, "/dicsasdsfdf[a]df[][][][[]asd[f][sda]f[xc]")
sheet.createPDF()

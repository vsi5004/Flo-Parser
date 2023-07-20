import xlsxwriter

title_str = "Main Title:"
author_str = "Author:"
format_str = "Format:"
pages_str = "Number of Pages:"
notes_str = "Notes:"
holdings_str = "Holdings Information"
filename = "webpage.txt"
output_file = "output.xlsx"

class Bibliography:
    def __init__(self, raw_data:list[str]) -> None:
        self.raw = raw_data
        self.author = None
        self.title = None
        self.format = None
        self.notes = ""
        self.pages = ""
        self.parse()

    def parse(self):
        counter = 0
        notes_start_index = 0
        notes_end_index = 0
        for line in self.raw:
            if title_str in line:
                self.title = line.lstrip(title_str).strip('\t').rstrip(".")
            elif author_str in line:
                self.author = line.lstrip(author_str).strip('\t').rstrip(".")
            elif format_str in line:
                self.format = line.lstrip(format_str).strip('\t')
            elif pages_str in line:
                self.pages = line.lstrip(pages_str).strip('\t')
            elif notes_str in line:
                notes_start_index = counter
            elif holdings_str in line:
                notes_end_index = counter
            counter+=1
        
        if notes_start_index > 0 and notes_end_index>notes_start_index:
            for i in range(notes_start_index, notes_end_index):
                self.notes+=self.raw[i].lstrip(notes_str).lstrip('\t').lstrip('\t')+" "

        if not self.author and " / " in self.title:
            self.author = self.title.split(" / ")[0]
        
        name = self.author.split(", ")
        if len(name)>1:
            name = name[1] +" "+name[0]
            if (" by "+name) in self.title:
                self.title = self.title.rstrip(" by "+name)

        if " / by" in self.title:
            self.title = self.title.split(" / ")[0]

        self.title = self.title.rstrip(" /")
        


with open(filename, encoding='utf-8') as my_file:
    lines = my_file.readlines()
    entries = []
    bib = None
    end_char = '\n'
    last_char = 'd'

    for line in lines:
        line = line.rstrip(end_char)
        if "Persistent link to this record" in line:
            if bib:
                entries.append(bib)
            bib=[]

        bib.append(line)

    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()
    worksheet.write(0,0,"Author")
    worksheet.write(0,1,"Title")
    worksheet.write(0,2,"Format")
    worksheet.write(0,3,"Pages")
    worksheet.write(0,4,"Notes")
    row = 1
    for entry in entries:
        bib = Bibliography(entry)
        worksheet.write(row,0,bib.author)
        worksheet.write(row,1,bib.title)
        worksheet.write(row,2,bib.format)
        worksheet.write(row,3,bib.pages)
        worksheet.write(row,4,bib.notes)
        row+=1

    workbook.close()

    print("Finished parsing data")
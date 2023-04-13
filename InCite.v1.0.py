import sys
from datetime import datetime
import nbib
import win32com.client as win32
# This used to be win32.constants.wdFieldCitation, but it's not working well. This might not be an issue with the compiled exe, but until then
#wdFieldCitation = win32.constants.wdFieldCitation
wdFieldCitation = 96


def add_citation(xml_string):

    # Select Microsoft Word
    word = win32.GetActiveObject("Word.Application")
    # Select the active document (not sure if this will work if multiple files are open, but one is in front? could specify a specific file in the future with:
    # doc = wordapp.Documents.Open(r"C:\\temp\\testing.docx")
    doc = word.ActiveDocument
    # Adds the formatted xml from the other functions to the document, which should also add it to the master list I think.
    doc.Bibliography.Sources.Add(xml_string)

def add_citation_tag(tags):
    # Select Microsoft Word
    word = win32.GetActiveObject("Word.Application")
    # Select the active document
    doc = word.ActiveDocument
    # Get where the cursor is
    cursor_pos = word.Selection.Range.Start

    # Look to see if the cursor is in a citation field
    # Pull out all fields 
    fields = doc.Fields
    current_field = None
    for field in fields:
        # Check if cursor is within the bounds of each field
        # # Note that there's 2 places of buffer for this, which may help with some finnicky stuff.
        if cursor_pos >= field.Result.Start-2 and cursor_pos <= field.Result.End+2:
            current_field = field
            break
    # If the cursor IS in a field, 
    if current_field is not None:
        # AND it's a citation field then this will happen
        if current_field.Type == wdFieldCitation:
            
            # Pull the already present tags, which have a '-' in them (although in the Source Manager it doesn't show them)
            current_code = current_field.Code.Text.split(' ')
            current_tags = [entry for entry in current_code if '-' in entry]

            # Only put in the new tag if it's not already in the existing list
            for tag in tags:
                if tag not in current_tags:
                    current_tags.append(tag)
            # Make a combo with the \m separator for input into a new field
            new_field_code = " \\m ".join(current_tags)     # Just lists the tags and Fields.Add puts in the rest

            #current_field.Code.Text = f"{current_code} \\m {tag}"
            # This doesn't work due to protections, that may be due to Office 365

            # Removes current_field since I can't figure out how to edit it
            current_field.Delete()
            # Redefines the cursor_pos since the removal of the previous field will change where the cursor is
            cursor_pos = word.Selection.Range.Start
            # Adds new combined citation field
            doc.Fields.Add(doc.Range(cursor_pos, cursor_pos), wdFieldCitation, Text=new_field_code, PreserveFormatting = True)

    else:
        if len(tags) > 1:
            tags_formatted = " \\m ".join(tags)
        else:
            tags_formatted = tags[0]
        # \\l 1033 just says means that it's in english, not really needed it seems
        # Adds citation field
        doc.Fields.Add(doc.Range(cursor_pos, cursor_pos), wdFieldCitation, Text=tags_formatted, PreserveFormatting = True)

    # Update the fields in the document, not sure how necessary it is, but it's here
    doc.Fields.Update()


def converting_and_citing(format, file):

    bib_dict = {}

    ##########################################################
    # Pulling data from file into a dictionary, bib_dict


    if format == 'ris':
        with open(file, 'r',encoding='utf-8') as ris_file:
            authors = []
            editors = []
            for line in ris_file:
                field = line[0:4].strip()
                value = line[6:].strip()
                if field == 'AU':
                    last = value.split(',')[0]
                    first_middle = value.split(',')[1].strip()
                    first = first_middle.split(' ')[0]
                    if len(first_middle.split(' ')) > 1:
                        middle = first_middle.split(' ')[1]
                        author = {'first_name':first,'Middle':middle,'last_name':last}
                    else:
                        author = {'first_name':first,'last_name':last}
                    authors.append(author)
                elif field == 'ED':
                    last = value.split(',')[0]
                    first_middle = value.split(',')[1].strip()
                    first = first_middle.split(' ')[0]
                    if len(first_middle.split(' ')) > 1:
                        middle = first_middle.split(' ')[1]
                        editor = {'first_name':first,'Middle':middle,'last_name':last}
                    else:
                        editor = {'first_name':first,'last_name':last}
                    editors.append(editor)
                elif field not in bib_dict:
                    bib_dict[field] = value
                else:
                    past_values = list(bib_dict[field])
                    bib_dict[field] = []
                    bib_dict[field].append(value)
                    bib_dict[field].extend(past_values)

            bib_dict['Author'] = authors
            bib_dict['Editor'] = editors
            if 'TI' not in bib_dict:
                if 'T1' in bib_dict:
                    bib_dict['TI'] = bib_dict['T1']


        ##############################
        # convert to windows xml string

        xml_string = '<b:Source>'
        # source type
        # JOUR = JournalArticle
        #xml_string += '<b:SourceType>'+bib_dict['TY']+'</b:SourceType>'
        xml_string += '<b:SourceType>'+'JournalArticle'+'</b:SourceType>'
        # tag
        tag = bib_dict['Author'][0]['last_name'][0:3].lower().strip()+bib_dict['PY'][-2:]+'-'+bib_dict['TI'][-5:].lower().strip()       # tags need to be unique!
        if tag in all_tags:
            add_citation_tag([tag])
            sys.exit()

        xml_string += '<b:Tag>'+tag+'</b:Tag>'
        # title
        xml_string += '<b:Title>'+bib_dict['TI']+'</b:Title>'
        # journal name
        xml_string += '<b:JournalName>'+bib_dict['JO']+'</b:JournalName>'
        # year
        xml_string += '<b:Year>'+bib_dict['PY']+'</b:Year>'
        # pages
        #xml_string += '<b:Pages>'+bib_dict['SN']+'</b:Pages>'
        # authors and editors
        xml_string += '<b:Author>'
        
        # authors
        xml_string += '<b:Author><b:NameList>'
        for author_dict in bib_dict['Author']:
            if 'Middle' in author_dict:
                xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First><b:Middle>'+author_dict['Middle']+'</b:Middle></b:Person>'
            else:
                xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First></b:Person>'
        xml_string += '</b:NameList></b:Author>'

        # editors
        if bib_dict['Editor'] != []:
            xml_string += '<b:Editor><b:NameList>'
            for editor_dict in bib_dict['Editor']:
                if 'Middle' in author_dict:
                    xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First><b:Middle>'+author_dict['Middle']+'</b:Middle></b:Person>'
                else:
                    xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First></b:Person>'
            xml_string += '</b:NameList></b:Editor>'
        else:
            xml_string += '</b:Author>'

        #city
        if 'CY' in bib_dict:
            xml_string += '<b:City>'+bib_dict['CY']+'</b:City>'
        # month published
        if 'DA' in bib_dict:
            month = bib_dict['DA'].split('/')[1]
            day = bib_dict['DA'].split('/')[2]
            xml_string += '<b:Month>'+month+'</b:Month>'
            xml_string += '<b:Day>'+day+'</b:Day>'
        # publisher
        if 'PB' in bib_dict:
            xml_string += '<b:Publisher>'+bib_dict['PB']+'</b:Publisher>'
        # volume
        if 'VL' in bib_dict:
            xml_string += '<b:Volume>'+bib_dict['VL']+'</b:Volume>'
        # issue
        if 'IS' in bib_dict:
            xml_string += '<b:Issue>'+bib_dict['IS']+'</b:Issue>'
        # short title
        if 'ST' in bib_dict:
            xml_string += '<b:ShortTitle>'+bib_dict['ST']+'</b:ShortTitle>'
        # standard number
        if 'VO' in bib_dict:
            xml_string += '<b:StandardNumber>'+bib_dict['VO']+'</b:StandardNumber>'
        # no comments
        # accessed date
        if 'Y2' in bib_dict:
            year = bib_dict['Y2'].split('/')[0]
            month = bib_dict['Y2'].split('/')[1]
            day = bib_dict['Y2'].split('/')[2]
            xml_string += '<b:YearAccessed>'+year+'</b:YearAccessed>'
            xml_string += '<b:MonthAccessed>'+month+'</b:MonthAccessed>'
            xml_string += '<b:DayAccessed>'+day+'</b:DayAccessed>'
        else:
            date = datetime.today().strftime('%Y-%m-%d').split('-')
            year = date[0]
            month = date[1]
            day = date[2]
            xml_string += '<b:YearAccessed>'+year+'</b:YearAccessed>'
            xml_string += '<b:MonthAccessed>'+month+'</b:MonthAccessed>'
            xml_string += '<b:DayAccessed>'+day+'</b:DayAccessed>'
        #URL
        if 'UR' in bib_dict:
            xml_string += '<b:URL>'+bib_dict['UR']+'</b:URL>'
        #DOI
        if 'DO' in bib_dict:
            xml_string += '<b:DOI>'+bib_dict['DO']+'</b:DOI>'

        # end xml
        xml_string += '</b:Source>'

        return(tag, xml_string)



    if format == 'nbib':

        ##########################################################
        # Pulling data from file into a dictionary, bib_dict
        # Huge thank you to Karl Holub!! I hope there are more packages like these in the future https://github.com/holub008/nbib
        bib_dict = nbib.read_file(file)[0]

        ##############################
        # convert to windows xml string

        xml_string = '<b:Source>'
        # source type
        # JOUR = JournalArticle
        #xml_string += '<b:SourceType>'+bib_dict['TY']+'</b:SourceType>'
        xml_string += '<b:SourceType>'+bib_dict['publication_types'][0].replace(' ','')+'</b:SourceType>'
        # tag
        tag = bib_dict['authors'][0]['last_name'][0:3]+bib_dict['publication_date'].split(' ')[0][-2:]+'-'+bib_dict['title'][-5:]      # tags need to be unique!
        if tag in all_tags:
            add_citation_tag([tag])
            sys.exit()
        xml_string += '<b:Tag>'+tag+'</b:Tag>'
        # skipping guid since I don't think it's necessary
        #xml_string += '<b:Guid>{1A809168-CAE0-420A-9D2F-5EF9815DCE7F}</b:Guid>'
        # title
        xml_string += '<b:Title>'+bib_dict['title']+'</b:Title>'
        # journal name
        xml_string += '<b:JournalName>'+bib_dict['journal']+'</b:JournalName>'
        # year
        pub_date = bib_dict['publication_date']
        year = pub_date.split(' ')[0]
        month = pub_date.split(' ')[1]
        xml_string += '<b:Year>'+year+'</b:Year>'
        # pages
        xml_string += '<b:Pages>'+bib_dict['pages']+'</b:Pages>'
        # authors and editors
        xml_string += '<b:Author>'
        
        # authors
        xml_string += '<b:Author><b:NameList>'
        for author_dict in bib_dict['authors']:
            if 'middle_name' in author_dict:
                xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First><b:Middle>'+author_dict['middle_name']+'</b:Middle></b:Person>'
            else:
                xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First></b:Person>'
        xml_string += '</b:NameList></b:Author>'

        # editors
        if 'editors' in bib_dict:
            xml_string += '<b:Editor><b:NameList>'
            for editor_dict in bib_dict['Editor']:
                if 'middle_name' in author_dict:
                    xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First><b:Middle>'+author_dict['middle_name']+'</b:Middle></b:Person>'
                else:
                    xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First></b:Person>'
            xml_string += '</b:NameList></b:Editor>'
        else:
            xml_string += '</b:Author>'

        #city
        if 'place_of_publication' in bib_dict:
            xml_string += '<b:City>'+bib_dict['place_of_publication']+'</b:City>'
        # month published
        if 'DA' in bib_dict:
            month = bib_dict['DA'].split('/')[1]
            day = bib_dict['DA'].split('/')[2]
            xml_string += '<b:Month>'+month+'</b:Month>'
            xml_string += '<b:Day>'+day+'</b:Day>'
        # publisher
        if 'PB' in bib_dict:
            xml_string += '<b:Publisher>'+bib_dict['PB']+'</b:Publisher>'
        # volume
        if 'journal_volume' in bib_dict:
            xml_string += '<b:Volume>'+bib_dict['journal_volume']+'</b:Volume>'
        # issue
        if 'journal_issue' in bib_dict:
            xml_string += '<b:Issue>'+bib_dict['journal_issue']+'</b:Issue>'
        # short title
        if 'short_title' in bib_dict:
            xml_string += '<b:ShortTitle>'+bib_dict['short_title']+'</b:ShortTitle>'
        # standard number
        if 'pubmed_id' in bib_dict:
            xml_string += '<b:StandardNumber>'+str(bib_dict['pubmed_id'])+'</b:StandardNumber>'
        # no comments
        # accessed date
        if 'Y2' in bib_dict:
            year = bib_dict['Y2'].split('/')[0]
            month = bib_dict['Y2'].split('/')[1]
            day = bib_dict['Y2'].split('/')[2]
            xml_string += '<b:YearAccessed>'+year+'</b:YearAccessed>'
            xml_string += '<b:MonthAccessed>'+month+'</b:MonthAccessed>'
            xml_string += '<b:DayAccessed>'+day+'</b:DayAccessed>'
        else:
            date = datetime.today().strftime('%Y-%m-%d').split('-')
            year = date[0]
            month = date[1]
            day = date[2]
            xml_string += '<b:YearAccessed>'+year+'</b:YearAccessed>'
            xml_string += '<b:MonthAccessed>'+month+'</b:MonthAccessed>'
            xml_string += '<b:DayAccessed>'+day+'</b:DayAccessed>'
        #URL
        if 'url' in bib_dict:
            xml_string += '<b:URL>'+bib_dict['url']+'</b:URL>'
        #DOI
        if 'doi' in bib_dict:
            xml_string += '<b:DOI>'+bib_dict['doi']+'</b:DOI>'

        # end xml
        xml_string += '</b:Source>'

        return(tag, xml_string)


    if format == 'enw':

        ##########################################################
        # Pulling data from file into a dictionary, bib_dict

        with open(file, 'r',encoding='utf-8') as enw_file:
            authors = []
            for line in enw_file:
                field = line.split(' ')[0]
                value = ' '.join(line.split(' ')[1:]).strip()
                if field == '%A':
                    last = value.split(',')[0]
                    first_middle = value.split(', ')[1]
                    if ' ' in first_middle:
                        first = first_middle.split(' ')[0]
                        middle = first_middle.split(' ')[1]
                        author = {'first_name':first,'Middle':middle,'last_name':last}
                    else:
                        first = first_middle
                        author = {'first_name':first,'last_name':last}
                        
                    authors.append(author)
                elif field not in bib_dict:
                    bib_dict[field] = value
                else:
                    past_values = list(bib_dict[field])
                    bib_dict[field] = []
                    bib_dict[field].append(value)
                    bib_dict[field].extend(past_values)

            bib_dict['Author'] = authors

        # Just a note that %@ is for the electronic ISBN and
        # %X is for the abstract, although it seems these aren't used in the windows citation manager
        # for journals


        ##############################
        # convert to windows xml string

        xml_string = '<b:Source>'
        # source type
        # JOUR = JournalArticle
        #xml_string += '<b:SourceType>'+bib_dict['TY']+'</b:SourceType>'
        xml_string += '<b:SourceType>'+'JournalArticle'+'</b:SourceType>'
        # tag
        tag = bib_dict['Author'][0]['last_name'][0:3]+bib_dict['%D'][-2:]+'-'+bib_dict['%T'][-5:]   # tags need to be unique
        while tag in all_tags:
            bib_dict['Author'][0]['last_name'][0:3]+bib_dict['%D'][-2:]+'-'+bib_dict['%T'][-5:]
        xml_string += '<b:Tag>'+tag+'</b:Tag>'
        # title
        xml_string += '<b:Title>'+bib_dict['%T']+'</b:Title>'
        # journal name
        xml_string += '<b:JournalName>'+bib_dict['%J']+'</b:JournalName>'
        # year
        xml_string += '<b:Year>'+bib_dict['%D']+'</b:Year>'
        # pages
        xml_string += '<b:Pages>'+bib_dict['%P']+'</b:Pages>'
        # authors and editors
        xml_string += '<b:Author>'
        
        # authors
        xml_string += '<b:Author><b:NameList>'
        for author_dict in bib_dict['Author']:
            if 'Middle' in author_dict:
                xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First><b:Middle>'+author_dict['Middle']+'</b:Middle></b:Person>'
            else:
                xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First></b:Person>'
        xml_string += '</b:NameList></b:Author>'

        # editors, seems no editors in this format
        if 'Editor' in bib_dict:
            xml_string += '<b:Editor><b:NameList>'
            for editor_dict in bib_dict['Editor']:
                if 'Middle' in author_dict:
                    xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First><b:Middle>'+author_dict['Middle']+'</b:Middle></b:Person>'
                else:
                    xml_string += '<b:Person><b:Last>'+author_dict['last_name']+'</b:Last><b:First>'+author_dict['first_name']+'</b:First></b:Person>'
            xml_string += '</b:NameList></b:Editor>'
        else:
            xml_string += '</b:Author>'

        #city
        if 'CY' in bib_dict:
            xml_string += '<b:City>'+bib_dict['CY']+'</b:City>'
        # month published
        if '%8' in bib_dict:
            month = bib_dict['%8'].split('/')[1]
            day = bib_dict['%8'].split('/')[2]
            xml_string += '<b:Month>'+month+'</b:Month>'
            xml_string += '<b:Day>'+day+'</b:Day>'
        # publisher
        if '%I' in bib_dict:
            xml_string += '<b:Publisher>'+bib_dict['%I']+'</b:Publisher>'
        # volume
        if '%V' in bib_dict:
            xml_string += '<b:Volume>'+bib_dict['%V']+'</b:Volume>'
        # issue
        if '%N' in bib_dict:
            xml_string += '<b:Issue>'+bib_dict['%N']+'</b:Issue>'
        # short title
        if 'ST' in bib_dict:
            xml_string += '<b:ShortTitle>'+bib_dict['ST']+'</b:ShortTitle>'
        # standard number
        if '%M' in bib_dict:
            xml_string += '<b:StandardNumber>'+bib_dict['%M']+'</b:StandardNumber>'
    # no comments
        # accessed date
        if '%[' in bib_dict:
            month = bib_dict['%['].split('/')[0]
            day = bib_dict['%['].split('/')[1]
            year = bib_dict['%['].split('/')[2]
            xml_string += '<b:YearAccessed>'+year+'</b:YearAccessed>'
            xml_string += '<b:MonthAccessed>'+month+'</b:MonthAccessed>'
            xml_string += '<b:DayAccessed>'+day+'</b:DayAccessed>'
        else:
            date = datetime.today().strftime('%Y-%m-%d').split('-')
            year = date[0]
            month = date[1]
            day = date[2]
            xml_string += '<b:YearAccessed>'+year+'</b:YearAccessed>'
            xml_string += '<b:MonthAccessed>'+month+'</b:MonthAccessed>'
            xml_string += '<b:DayAccessed>'+day+'</b:DayAccessed>'
        #URL
        if '%R' in bib_dict:
            xml_string += '<b:URL>'+bib_dict['%R'].strip()+'</b:URL>'
        #DOI
        if '%U' in bib_dict:
            doi = bib_dict['%U'].strip()
            doi = doi.replace('doi:','')
            if 'doi.org/' not in doi:
                doi = 'https://doi.org/'+doi
            xml_string += '<b:DOI>'+doi+'</b:DOI>'

        # end xml
        xml_string += '</b:Source>'
        return(tag, xml_string)


##############################
#   MAIN
##############################

# Pull in the citation file

# Check if an input file was given
if len(sys.argv) > 1:

    # If so then set it and the file type
    file = sys.argv[1]
    extension = file.split('.')[-1]

    # Then check if the extension if good
    if extension in ['enw','nbib','ris']:

        # Try to connect with the open word document, throw an appropriate error if cannot
        try:
            word = win32.GetActiveObject("Word.Application")
            doc = word.ActiveDocument
        except:
            raise Exception("There doesn't appear to be an open word document. Make sure that a word document is open and that the cursor is at the location where you would like to insert the citation.")

        # Find how many sources are already in the document
        ref_count = doc.Bibliography.Sources.Count
        # Store the source tags from the document 
        all_tags = []
        for i in range(1,ref_count+1):  # It's base 1
            all_tags.append(doc.Bibliography.Sources(i).Tag)

        tag, xml_string = converting_and_citing(extension, file)
        if tag not in all_tags:
            add_citation(xml_string)
        add_citation_tag([tag])


##############################
#   MAIN for search GUI if no file
##############################

# If no input file is given then make the search GUI
else:

    import tkinter as tk
    from tkinter import ttk, Tk, Entry, PhotoImage, filedialog, messagebox, Listbox, END, INSERT
    import fuzzysearch
    import xml.etree.ElementTree as ET


    # Try to connect with the open word document, throw an appropriate error if cannot
    try:
        word = win32.GetActiveObject("Word.Application")
        doc = word.ActiveDocument
    except:
        raise Exception("There doesn't appear to be an open word document. Make sure that a word document is open and that the cursor is at the location where you would like to insert the citation.")

    # Find how many sources are already in the document
    ref_count = doc.Bibliography.Sources.Count
    # Store the source tags from the document 
    all_xml_sources = []
    for i in range(1,ref_count+1):  # It's base 1
        all_xml_sources.append(doc.Bibliography.Sources(i).XML.replace('xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"','xmlns:b="source"'))


    converted_list = []
    for xml_source in all_xml_sources:
        ET_tree = ET.fromstring(xml_source)
        title = ET_tree.find('{source}Title').text
        journal = ET_tree.find('{source}JournalName').text
        year = ET_tree.find('{source}Year').text
        tag = ET_tree.find('{source}Tag').text
        authors = []
        for author in ET_tree.find('{source}Author').iter('{source}Person'):
            last_name = author.find('{source}Last').text
            first_name = author.find('{source}First').text
            full_name = f"{last_name}, {first_name[0]}"     # This is just using first name's initial
            #full_name = f"{last_name}, {first_name}"     # This is using the full first name
            authors.append(full_name)
        author_str = '; '.join(authors)
        converted_list.append((author_str,year,journal,title,tag))

    def search(search_string,event=None):
        output_list = []
        for source in converted_list:
            authors = source[0]
            list_string = ''.join([source[3],source[1],source[2],authors])

            results = fuzzysearch.find_near_matches(search_string.lower(), list_string.lower(), max_l_dist=2)
            if len(results) > 0:
                top_result = results[0]
                if top_result.dist < 3:
                    new_source = list(source) + [top_result.dist]
                    output_list.append(new_source)
        output_list = sorted(output_list, key=lambda source: source[-1])
        tree.delete(*tree.get_children())
        for citation in output_list:
            tree.insert('', tk.END, values=citation)
        first_item = tree.get_children()[0]
        tree.selection_set(first_item)
        

    
    def submit(event=None):
        chosen_tags = []
        #curItem = tree.focus()
        for selected_item in tree.selection():
            item = tree.item(selected_item)
            chosen_tags.append(item['values'][4])
        add_citation_tag(chosen_tags)


    def treeview_sort_column(tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # reverse sort next time
        tv.heading(col, command=lambda: \
                treeview_sort_column(tv, col, not reverse))

    root = Tk()
    root.title('InCite Current Citation Search')

    main_frame = ttk.Frame(root)
    main_frame.pack(fill='both', expand=True)

    search_label_var = ttk.Label(main_frame, text='Search:')
    search_label_var.grid(column=0, row=0, padx=5,
                        pady=5, sticky='ew')
    
    search_entry_var = ttk.Entry(main_frame, width='20')
    search_entry_var.insert(0, '')
    search_entry_var.grid(column=1, row=0,
                        padx=5, pady=5, sticky='ew')
    search_entry_var.focus()

    search_label2_var = ttk.Label(main_frame, text='(<Enter>)')
    search_label2_var.grid(column=3, row=0, padx=5,
                        pady=5, sticky='ew')
    
    columns = ('Authors','Year','Journal','Title')
    tree = ttk.Treeview(main_frame, columns=columns,show='headings')
    tree.column('Authors', width=300,minwidth=50)
    tree.column('Year', width=50,minwidth=50)
    tree.column('Journal', width=200,minwidth=50)
    tree.column('Title', width=400,minwidth=50)

    for citation in converted_list:
        tree.insert('', tk.END, values=citation)
    for col in columns:
        tree.heading(col, text=col, command=lambda _col=col: \
                     treeview_sort_column(tree, _col, False))
    tree.grid(column=0, row=2,columnspan=10,
                         padx=5, pady=5, sticky='ew')

    search_button = ttk.Button(main_frame, text='Search', command=lambda: search(search_entry_var.get()))
    search_button.grid(column=2, row=0, sticky='se', padx=5, pady=5)

    submit_button = ttk.Button(main_frame, text='Insert Citation', command=submit)
    submit_button.grid(column=4, row=0, sticky='se', padx=5, pady=5)

    search_label2_var = ttk.Label(main_frame, text='(<Control+Enter>)')
    search_label2_var.grid(column=5, row=0, padx=5,
                        pady=5, sticky='ew')


    # makes it so that hitting 'Enter' triggers the search button.
    root.bind('<Return>', lambda _: search(search_entry_var.get()))
    # and hitting 'Ctrl+Enter' triggers the submittion
    root.bind('<Control-Return>', submit)

    # This is what constantly updates the window.
    root.mainloop()
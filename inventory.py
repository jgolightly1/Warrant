import sys
import FreeSimpleGUI as sg
from pathlib import Path
from docxtpl import DocxTemplate
import datetime

today=datetime.datetime.today()

document_path = Path(__file__).parent / "Inventory_Template.docx"
doc = DocxTemplate(document_path)

# ----------- Set the Theme -----------
sg.theme('DarkGrey5')

# ----------- Set Certain Keys for DOCXTPL -----------
context = {
    'RESIDENTIAL':'RESIDENTIAL', 'CSAM':'CSAM', 'GANG':'GANG', 'NARCOTICS':'NARCOTICS', 'MURDER':'MURDER', 'BAC':'BAC',
    'ATT':'ATT', 'VERIZON':'VERIZON', 'INFO':'INFO'
    }

# ----------- Create Different Columns for Different Warrants -----------
col1 = sg.Column([[sg.Frame('Case Identifiers:',
                            [[sg.Column([[sg.Text('What would you like to name this file?'), sg.Input(key='REF')],
                            [sg.Text('What is your assigned Case Number?'), sg.Input(key='CASE')],
                            ])]])]])

col2 = sg.Push(), sg.Column([[sg.Frame('Information:',[[sg.Text(), sg.Column([
    [sg.Text('What County are you in?')],
    [sg.Input(key='COUNTY')],
    [sg.Text('What is your Name?')],
    [sg.Input(key='NAME')],
    [sg.Text('Who do you work for?')],
    [sg.Input(key='AGENCY')],
    [sg.Text('What Date was the Warrant served? (DD/MM/YYYY)')],
    [sg.Input(key='DATE')]
    ])]])]]), sg.Push()

col0 = sg.Column([[sg.Frame('Actions:',
                            [[sg.Column([[sg.Button('Save'), sg.Exit()]])]])]])

# ----------- Create Different Windows -----------
def make_welcome():        
    layout = [ 
    [sg.Text('Sikeston Department of Public Safety', font=('helvetica', '14'), pad=(100,10))],
    [sg.Text('Criminal Investigations Unit', font=('helvetica', '13'))],
    [sg.Image('Patch.png', pad=(50,10))],
    [sg.Text('Return and Inventory', font=('helvetica', '12'))],
    [sg.Button('Begin', font=('helvetica', '12'), pad=(10,10))], 
    [sg.Text('Search Warrant Creation Tool', font=('helvetica', '12'))],
    [sg.Text('Version: 1.0', font=('helvetica', '12'))],
    [sg.Button('Exit', font=('helvetica', '12'), pad=(10,10))]]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, element_justification='c', finalize=True)

def make_begin():
    layout = [sg.Push(), col1, sg.Push()], [col2], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

window1, window2, window3 = make_welcome(), None, None

# ----------- Event Loop -----------
while True:
    window, event, context = sg.read_all_windows()
    if window == window2 and event in (sg.WIN_CLOSED, 'Exit'):
        break

    if window == window1:
        if event == 'Begin':
            window1.close()
            window2 = make_begin()


    if event == "Save":
        # Add calculated fields to our dict

        context["TODAY"] = today.strftime("%m/%d/%Y")

        # Render the template, save new word document & inform user
        doc.render(context)

        output_path = Path(__file__).parent / 'Complete/' / f"{context['REF']}-SW-Inventory.docx"
#        output_path = Path(__file__).parent / f"{context['REF']}-SW-Inventory.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        break

    if event == 'Exit':
        break

window.close()

import sys
import PySimpleGUI as sg
from pathlib import Path
from docxtpl import DocxTemplate
import datetime

today=datetime.datetime.today()

document_path = Path(__file__).parent / "CDR_Template.docx"
doc = DocxTemplate(document_path)

# ----------- Set the Theme -----------
sg.theme('DarkGrey5')

# ----------- Set Certain Keys for DOCXTPL -----------
context = {
    'RESIDENTIAL':'RESIDENTIAL', 'CSAM':'CSAM', 'GANG':'GANG', 'NARCOTICS':'NARCOTICS', 'MURDER':'MURDER', 'BAC':'BAC',
    'ATT':'ATT', 'VERIZON':'VERIZON', 'TMOBILE':'TMOBILE', 'USCELLULAR':'USCELLULAR', 'PING':'PING', 'PRTT':'PRTT'
    }

# ----------- Create Different Columns for Different Warrants -----------
col1 = sg.Column([[sg.Frame('Case Identifiers:',
                            [[sg.Column([[sg.Text('What would you like to name this file?'), sg.Input(key='REF')],
                            [sg.Text('What is your assigned Case Number?'), sg.Input(key='CASE')],
                            ])]])]])

col2 = sg.Column([[sg.Frame('Information and Background:', [[sg.Text(), sg.Column([
    [sg.Text('What County are you in?')],
    [sg.Input(key='COUNTY')],
    [sg.Text('What is your Name?')],
    [sg.Input(key='NAME')],
    [sg.Text('Who do you work for?')],
    [sg.Input(key='AGENCY')],
    [sg.Text('What is your Rank?')],
    [sg.Input(key='RANK')],
    [sg.Text('How long have you been a police officer (in years)?')],
    [sg.Input(key='YEARS')],
    [sg.Text('What crime are you investigating?')],
    [sg.Input(key='CRIME')],
    [sg.Text("What's the statute number?")],
    [sg.Input(key='RSMO')],
    [sg.Text('What day (1,2,3) did this occur?')],
    [sg.Input(key='DAY')],
    [sg.Text('What Month?')],
    [sg.Input(key='MONTH')],
    [sg.Text('What Year?')],
    [sg.Input(key='YEAR')],
    ])]])]])

col3 = sg.Column([
    [sg.Frame('Information:', [[sg.Text(), sg.Column([
    [sg.Checkbox('Check this box to add Phone Ping', key='PING')],
    [sg.Text('Target Telephone Number:')],
    [sg.Input(key='PHONENUMBER')],
    [sg.Text('Historical records usually request 3 months to establish patter of life')],
    [sg.Text('Start Date:')],
    [sg.Text('Example: 01/01/2023 00:00 CST')],
    [sg.Input(key='STARTDATE')],
    [sg.Text('End Date:')],
    [sg.Text('Example: 01/31/2023 23:59 CST')],
    [sg.Input(key='ENDDATE')],
    [sg.Text('What is your email address:')],
    [sg.Input(key='EMAIL')],
    [sg.Text('Who is the Target Number cell provider?')],
    [sg.Push(), sg.Text('Does your warrant require any special language?'), sg.Push()],
                            [sg.Checkbox('AT&T', key='ATT'), sg.Checkbox('T-Mobile', key='TMOBILE'), sg.Checkbox('US Cellular', key='USCELLULAR'),
                             sg.Checkbox('Verizon', key='VERIZON')],
    ])]])]])

col4 = sg.Column([
    [sg.Frame('Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Pen Register and Trap and Trace', 'radio1', key='PRTT', default=True)],
    [sg.Checkbox('Check this box to add Phone Ping', key='PING')],
    [sg.Text('Target Telephone Number:')],
    [sg.Input(key='PHONENUMBER')],
    [sg.Text('Historical records usually request 3 months to establish patter of life')],
    [sg.Text('Start Date:')],
    [sg.Text('Example: 01/01/2023 00:00 CST')],
    [sg.Input(key='STARTDATE')],
    [sg.Text('End Date:')],
    [sg.Text('Example: 01/31/2023 23:59 CST')],
    [sg.Input(key='ENDDATE')],
    [sg.Text('What is your email address:')],
    [sg.Input(key='EMAIL')],
    [sg.Push(), sg.Text('Does your warrant require any special language?'), sg.Push()],
                            [sg.Checkbox('AT&T', key='ATT'), sg.Checkbox('T-Mobile', key='TMOBILE'), sg.Checkbox('US Cellular', key='USCELLULAR'),
                             sg.Checkbox('Verizon', key='VERIZON')],
    ])]])]])

col0 = sg.Column([[sg.Frame('Actions:',
                            [[sg.Column([[sg.Button('Save'), sg.Exit()]])]])]])

# ----------- Create Different Windows -----------
def make_welcome():        
    layout = [ 
    [sg.Text('Sikeston Department of Public Safety', font=('helvetica', '14'), pad=(100,10))],
    [sg.Text('Criminal Investigations Unit', font=('helvetica', '13'))],
    [sg.Image('Patch.png', pad=(50,10))],
    [sg.Text('Electronic Communication Warrants', font=('helvetica', '12'))],
    [sg.Button('Call Detail Records', font=('helvetica', '12'), pad=(10,10)),
    sg.Button('Pen Register', font=('helvetica', '12'), pad=(10,10))],
    [sg.Text('Search Warrant Creation Tool', font=('helvetica', '12'))],
    [sg.Text('Version: 1.0', font=('helvetica', '12'))],
    [sg.Button('Exit', font=('helvetica', '12'), pad=(10,10))]]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, element_justification='c', finalize=True)

def make_cdr():
    layout = [sg.Push(), col1, sg.Push()], [col2, col3], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_prtt():
    layout = [sg.Push(), col1, sg.Push()], [col2, col4], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

window1, window2, window3 = make_welcome(), None, None

# ----------- Event Loop -----------
while True:
    window, event, context = sg.read_all_windows()
    if window == window2 and event in (sg.WIN_CLOSED, 'Exit'):
        break

    if window == window1:
        if event == 'Call Detail Records':
            window1.close()
            window2 = make_cdr()
                
        if event == 'Pen Register':
            window1.close()
            window2 = make_prtt()          

    if event == "Save":
        # Add calculated fields to our dict

        context["TODAY"] = today.strftime("%m/%d/%Y")

        # Render the template, save new word document & inform user
        doc.render(context)

        output_path = Path(__file__).parent / 'Complete/' / f"{context['REF']}-Search Warrant.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        break

    if event == 'Exit':
        break

window.close()
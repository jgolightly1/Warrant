import sys
import PySimpleGUI as sg
from pathlib import Path
from docxtpl import DocxTemplate
import datetime

today=datetime.datetime.today()

document_path = Path(__file__).parent / "Social_Template.docx"
doc = DocxTemplate(document_path)

# ----------- Set the Theme -----------
sg.theme('DarkGrey5')

# ----------- Set Certain Keys for DOCXTPL -----------
context = {
    'APPLE':'APPLE', 'GOOGLE':'GOOGLE', 'FACEBOOK':'FACEBOOK', 'FBPING':'FBPING', 'SNAP':'SNAP', 'SNAPPRTT':'SNAPPRTT', 'INSTA':'INSTA', 'INSTAPING':'INSTAPING', 'PING':'PING'
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
    [sg.Radio('Facebook', 'radio1', key='FACEBOOK', default=True)],
    [sg.Checkbox('Check this box to add Account Ping (Real Time Location)', key='FBPING')],
    [sg.Text('What is the Facebook ID or Account Name:')],
    [sg.Input(key='ACCOUNT')],
    [sg.Text('Historical records usually request 3 months to establish patter of life')],
    [sg.Text('Start Date:')],
    [sg.Text('Example: 01/01/2023 00:00 CST')],
    [sg.Input(key='STARTDATE')],
    [sg.Text('End Date:')],
    [sg.Text('Example: 01/31/2023 23:59 CST')],
    [sg.Input(key='ENDDATE')],
    ])]])]])

col4 = sg.Column([
    [sg.Frame('Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Snapchat', 'radio1', key='SNAP', default=True)],
    [sg.Checkbox('Check this box to add Pen Register / Trap and Trace', key='SNAPPRTT')],
    [sg.Text('What is the Snapchat Account Information:')],
    [sg.Input(key='ACCOUNT')],
    [sg.Text('Historical records usually request 3 months to establish patter of life')],
    [sg.Text('Start Date:')],
    [sg.Text('Example: 01/01/2023 00:00 CST')],
    [sg.Input(key='STARTDATE')],
    [sg.Text('End Date:')],
    [sg.Text('Example: 01/31/2023 23:59 CST')],
    [sg.Input(key='ENDDATE')],
    ])]])]])

col5 = sg.Column([
    [sg.Frame('Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Instagram', 'radio1', key='INSTA', default=True)],
    [sg.Checkbox('Check this box to add Account Ping (Real Time Location)', key='INSTAPING')],
    [sg.Text('What is the Instagram Account Information:')],
    [sg.Input(key='ACCOUNT')],
    [sg.Text('Historical records usually request 3 months to establish patter of life')],
    [sg.Text('Start Date:')],
    [sg.Text('Example: 01/01/2023 00:00 CST')],
    [sg.Input(key='STARTDATE')],
    [sg.Text('End Date:')],
    [sg.Text('Example: 01/31/2023 23:59 CST')],
    [sg.Input(key='ENDDATE')],
    ])]])]])

col6 = sg.Column([
    [sg.Frame('Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Apple', 'radio1', key='APPLE', default=True)],
    [sg.Text('What is the Apple ID:')],
    [sg.Input(key='ACCOUNT')],
    [sg.Text('Historical records usually request 3 months to establish patter of life')],
    [sg.Text('Start Date:')],
    [sg.Text('Example: 01/01/2023 00:00 CST')],
    [sg.Input(key='STARTDATE')],
    [sg.Text('End Date:')],
    [sg.Text('Example: 01/31/2023 23:59 CST')],
    [sg.Input(key='ENDDATE')],
    ])]])]])

col7 = sg.Column([
    [sg.Frame('Information:', [[sg.Text(), sg.Column([
    [sg.Radio('Google', 'radio1', key='GOOGLE', default=True)],
    [sg.Text('What is the Google Account Information:')],
    [sg.Input(key='ACCOUNT')],
    [sg.Text('Historical records usually request 3 months to establish patter of life')],
    [sg.Text('Start Date:')],
    [sg.Text('Example: 01/01/2023 00:00 CST')],
    [sg.Input(key='STARTDATE')],
    [sg.Text('End Date:')],
    [sg.Text('Example: 01/31/2023 23:59 CST')],
    [sg.Input(key='ENDDATE')],
    ])]])]])

col0 = sg.Column([[sg.Frame('Actions:',
                            [[sg.Column([[sg.Button('Save'), sg.Exit()]])]])]])

# ----------- Create Different Windows -----------
def make_welcome():        
    layout = [ 
    [sg.Text('Sikeston Department of Public Safety', font=('helvetica', '14'), pad=(100,10))],
    [sg.Text('Criminal Investigations Unit', font=('helvetica', '13'))],
    [sg.Image('Patch.png', pad=(50,10))],
    [sg.Text('Social Media Warrants', font=('helvetica', '12'))],
    [sg.Button('Apple', font=('helvetica', '12'), pad=(10,10)),
    sg.Button('Google', font=('helvetica', '12'), pad=(10,10)),
    sg.Button('Facebook', font=('helvetica', '12'), pad=(10,10)),
    sg.Button('Instagram', font=('helvetica', '12'), pad=(10,10)),
    sg.Button('Snapchat', font=('helvetica', '12'), pad=(10,10))],
    [sg.Text('Search Warrant Creation Tool', font=('helvetica', '12'))],
    [sg.Text('Version: 1.0', font=('helvetica', '12'))],
    [sg.Button('Exit', font=('helvetica', '12'), pad=(10,10))]]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, element_justification='c', finalize=True)

def make_facebook():
    layout = [sg.Push(), col1, sg.Push()], [col2, col3], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_snapchat():
    layout = [sg.Push(), col1, sg.Push()], [col2, col4], [sg.Push(), col0, sg.Push()]   

    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_instagram():
    layout = [sg.Push(), col1, sg.Push()], [col2, col5], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_apple():
    layout = [sg.Push(), col1, sg.Push()], [col2, col6], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

def make_google():
    layout = [sg.Push(), col1, sg.Push()], [col2, col7], [sg.Push(), col0, sg.Push()]
    
    return sg.Window('UNCLASSIFIDED // LAW ENFORCEMENT SENSITIVE', layout, font=('helvetica', '12'), element_padding=(5,5), finalize=True)

window1, window2, window3 = make_welcome(), None, None

# ----------- Event Loop -----------
while True:
    window, event, context = sg.read_all_windows()
    if window == window2 and event in (sg.WIN_CLOSED, 'Exit'):
        break

    if window == window1:
        if event == 'Facebook':
            window1.close()
            window2 = make_facebook() 

        if event == 'Instagram':
            window1.close()
            window2 = make_instagram()         

        if event == 'Apple':
            window1.close()
            window2 = make_apple() 

        if event == 'Google':
            window1.close()
            window2 = make_google() 

        if event == 'Snapchat':
            window1.close()
            window2 = make_snapchat() 

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
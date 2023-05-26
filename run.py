import PySimpleGUI as sg
import xml.etree.ElementTree as ET
import shutil
import zipfile
import os
import aspose.words as aw
import webbrowser
import uuid
import traceback
from collections import defaultdict

def process_question(q):
    """
    Processes an individual question node and extracts the necessary information.
    """
    question = {}
    for child in q:
        if child.tag == "material":
            question["text"] = child.find("mattext").text
        elif child.tag == "response_str":
            question["ftblist"] = [child.find("render_fib").find("response_label").attrib["ident"]]
    return question

sg.theme('reddit')
xmlFiles = []
ListOfQuiz = []

layout = [
    [sg.Text("NYPLMS Quiz File Generator",
             font='Courier 20  bold ')],
    [sg.Text("")],
    [sg.Text("Instruction:", font='5')],
    [sg.Text("Make sure the question types only contain the following:")],
    [sg.Text("1. True/False")],
    [sg.Text("2. Multiple Choice")],
    [sg.Text("3. Multi-Select")],
    [sg.Text("4. Short Answer")],
    [sg.Text("5. Long Answer")],
    [sg.Text("6. Multi-Short Answer")],
    [sg.Text("7. Fill in the Blanks")],
    [sg.Text(" ")],
    [sg.Text("Notes: ", text_color='red', font='5')],
    [sg.Text("Be sure that the quizzes from NYPLMS are exported in Brightspace Package format!")],
    [sg.Text(" ")],
    [sg.Text("Upload export file here:       "), sg.Input(key="-IN-"), sg.FileBrowse()],
    [sg.Text(" ")],
    [sg.Text("Enter filename to be saved: "), sg.Input(key="-NAME-")],
    [sg.Text(" ")],
    [sg.Button('User Manual')],
    [sg.Button("      Convert      "), sg.Exit()]

]

window=sg.Window("NYPLMS Export File Converter",layout,icon='LOGO.ico')

while True:
    event,values=window.read()

   
    if event == "      Convert      ":

        if values['Browse'] =="":
            sg.popup("Please Select a file to be coverted!")
            break

        try:

            filePath = values['Browse']

            # Generate a UUID
            unite_id = uuid.uuid4()

            # Convert the UUID to a string
            unite_id_str = str(unite_id)

            directoryPath=os.path.dirname(filePath)
            # Create a temporary directory for extracted files
            tempFileDirec = directoryPath + r"\temp" + unite_id_str
            os.makedirs(tempFileDirec)

            # Extract the contents of the zip file to the temporary directory
            with zipfile.ZipFile(filePath, "r") as zip_ref:
                zip_ref.extractall(tempFileDirec)

           
            workingFolder=tempFileDirec+"/questiondb.xml"

            NewFileName = values["-NAME-"]
            if NewFileName =='':
                zip_file_name =  os.path.basename(values["Browse"])
                final_name = os.path.splitext(zip_file_name)[0]
                NewFileName = final_name
            
            file_html=open(tempFileDirec+"/"+NewFileName+".html","w")
            accepted_types = ["True/False", "Multiple Choice", "Multi-Select", "Short Answer", "Multi-Short Answer", "Fill in the Blanks", "Long Answer"]
            QnList=[]                        
            ftblist = []                    
            optionlist=[]
            ansIdList={}
            ansDict = defaultdict(list)
            optionsDict = defaultdict(list)
            Question=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
            optionCounter=0
            qnNumber=0
            NumOfOptionPerQnCounter=0

            tree = ET.parse(workingFolder)
            root = tree.getroot()

            file_html.write("<html>")
            file_html.write("<head>")
            file_html.write("<style>table{border: 1px solid black;}img{max-width:50%;}</style>")


            for a in root.findall("./objectbank/section"):
                title=a.attrib['title']
                file_html.write("<title>{}</title></head><body>".format(title))
                file_html.write("<h2>{}</h2>".format(title))

                        
            for z in root.findall('./objectbank/section/item'):
                QnList.append(z.attrib['ident'])
            NumberOfQuestion=len(QnList)

            for x in root.findall('./objectbank/section/item/presentation/flow/material'):#qn
                optionCounter=0
                questiontext = ""
                file_html.write("<hr>")
                for question in x.findall('.//mattext'):
                    questiontext += question.text
                file_html.write("<h3> Question {} </h3>".format(qnNumber +1))
                for unQnType in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/itemmetadata/qtimetadata/qti_metadatafield[2]'):
                    badQntype=unQnType.find('fieldentry').text
                    if badQntype not in accepted_types: 
                        file_html.write("<p style='color: red;'>Attention : Question Type {}  not recognised at Qn No. {} !</p><br>".format(badQntype, qnNumber+1))
                        print("qn not found")
                    else:                        
                        file_html.write(questiontext)
                        # print(questiontext)
                        continue
                #
                for xyz in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/itemmetadata/qtimetadata/qti_metadatafield[3]'):
                        qnmark_str = xyz.find('fieldentry').text
                file_html.write("<p>  [ {} Marks ]</p>".format(int(float(qnmark_str))))
                        
                loopCounter=0

                for y in root.findall('./objectbank/section/item/presentation/flow/response_lid/render_choice/flow_label/response_label'):
                    loopCounter=loopCounter+1
                    ansId=y.attrib["ident"]
                    for w in root.findall(f'./objectbank/section/item/presentation/flow/response_lid/render_choice/flow_label/response_label[@ident="{ansId}"]/flow_mat/material'):
                        option=w.find('mattext').text
                        ansIdList[ansId]=option
                    for abc in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/itemmetadata/qtimetadata/qti_metadatafield[3]'):
                        qnmark_str = abc.find('fieldentry').text
                        # print(int(float(qnmark_str)))
                    

                    

                    for u in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/itemmetadata/qtimetadata/qti_metadatafield[2]'):
                            qntype=u.find('fieldentry').text
#####################################################
                            if qntype == "True/False" or qntype=="Multiple Choice":
                                for p in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}'):
                                    score=p.find('setvar').text
                                    try:
                                        formattedscore=int(float(score))
                                    except:
                                        formattedscore=score
                                    #print(formattedscore)
                                    if type(formattedscore) is int and formattedscore > 0:
                                        for v in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}/conditionvar'):
                                            correct=v.find('varequal').text
                                            try:
                                                AnsToBePrinted=ansIdList[correct]
                                                if AnsToBePrinted.startswith("<p>"):
                                                    formattedAnsToBePrinted=AnsToBePrinted[3:-4]
                                                    # file_html.write("<p>[ {} Marks ]</p>".format(int(float(qnmark_str))))
                                                    # file_html.write("<br><strong>Answer : {}  </strong><br>".format(formattedAnsToBePrinted))
                                                    ansDict[qnNumber+1].append(formattedAnsToBePrinted)
                                                else: 
                                                    # file_html.write("<p>[ {} Marks ]</p>".format(int(float(qnmark_str))))
                                                    # file_html.write("<br><strong>Answer : {}   </strong><br>".format(AnsToBePrinted))
                                                    ansDict[qnNumber+1].append(AnsToBePrinted)

                                            except:
                                                pass
                                            

                                    if loopCounter == 10:
                                        break
                                
                            elif qntype == "Multi-Select":
                                for p in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}/setvar'):
                                    score=p.attrib['varname']
                                    if score == "D2L_Correct":
                                        for v in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}/conditionvar'):
                                            correct=v.find('varequal').text
                                            try:
                                                AnsToBePrinted=ansIdList[correct]
                                                #print(AnsToBePrinted)
                                                if AnsToBePrinted.startswith("<p>"):
                                                    formattedAnsToBePrinted=AnsToBePrinted[3:-4]
                                                    # file_html.write("<strong >Answer : {}  </strong><br>".format(formattedAnsToBePrinted))
                                                    ansDict[qnNumber+1].append(formattedAnsToBePrinted)
                                                else:
                                                    # file_html.write("<strong>Answer : {}   </strong><br>".format(AnsToBePrinted))
                                                    ansDict[qnNumber+1].append(AnsToBePrinted)

                                                
                                            except:
                                                pass
                                    if loopCounter == 10:
                                        break
                              

                            elif qntype == "Long Answer":
                                temp=False
                                for h in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/answer_key/answer_key_material/flow_mat/material'):
                                    LongAnswer=h.find('mattext').text
                                    temp=True
                                    if temp ==False:
                                        for h in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/answer_key/answer_key_material/flow_mat/flow_mat'):
                                            LongAnswer=h.find('mattext').text
                                            file_html.write("<p>[ {} Marks ]</p>".format(int(float(qnmark_str))))
                                            
                                    else:
                                        file_html.write("<p>[ {} Marks ]</p>".format(int(float(qnmark_str))))
                                        file_html.write("<br><strong>Answer : {}   </strong><br>".format(LongAnswer))
                                        break

                            else:
                                continue

                            break
                    if qntype == "Long Answer":
                        break


                    if qntype == "Short Answer" or qntype == "Multi-Short Answer" or qntype == "Fill in the Blanks":
                        for o in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition/conditionvar'):
                            try:
                                LongAnswer=o.find('varequal').text
                                file_html.write("<br><strong>Answer : {}   </strong><br>".format(LongAnswer))
                                # flow_element = tree.find('.//presentation/flow')
                                # content = ''.join([child.text.strip() if child.text is not None else '' for child in list(flow_element)]).replace('response','_____')
                                # print(content)
                                # print('flow' + flow_element)
                                # print('list'+ list(flow_element))
                            except:
                                pass
                        else:
                            break
#######################################################
                for y in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/presentation/flow/response_lid/render_choice/flow_label/response_label/flow_mat/material'):#option
                    # option=y.find('mattext').text
                    response_lid = root.find(".//response_lid")
                    response_labels = response_lid.findall(".//response_label")
                    choices = {}
                    for label in response_labels:
                        ident = label.get('ident')
                        text = label.find(".//mattext").text.strip()
                        choices[ident] = text
                    print(choices)
                    option = y.find('mattext').text
                    if option.startswith("<p>"):
                        formatted=option[3:]
                        # 
                        
                        optionsDict[qnNumber+1].append(formatted)

                    else:        
                        
                        optionsDict[qnNumber+1].append(option)                    
                for key, value in optionsDict.items():
                    for i in range(len(value)):
                        value[i] = value[i].replace('</p>', '')
                # for key in ansDict.keys():
                #     set1 = set(ansDict[key])
                #     set2 = set(optionsDict[key])

                #     if qnNumber+1 == key:
                #         for i in set1 & set2:
                #             file_html.write("<p>{}. {}*".format(Question[optionCounter],i))
                #             optionCounter += 1
                #         for wat in set2 - set1:
                #             file_html.write("<p>{}. {}".format(Question[optionCounter],wat))
                #             optionCounter += 1
                for key in ansDict.keys():
                    set1 = set(ansDict[key])
                    set2 = list(optionsDict[key])  # Convert set2 to a list

                    if qnNumber + 1 == key:
                        modified_set2 = []
                        for wat in set2:
                            if wat in set1:
                                modified_wat = "{}.* {}".format(Question[optionCounter], wat)
                                optionCounter += 1
                            else:
                                modified_wat = "{}. {}".format(Question[optionCounter], wat)
                                optionCounter += 1
                            modified_set2.append(modified_wat)

                        for wat in modified_set2:
                            file_html.write("<p>{}".format(wat))
                        # for key in set2:
                            # if set2-set1[key] == key:
                            #     print(key)
                            # else:
                            #     print(key, 1111)
                else:
                    if QnList[qnNumber]==QnList[-1]:
                        break
                    else:
                        qnNumber=qnNumber+1

                for rand in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/presentation/flow/response_lid'):
                    randcheck = rand.find('render_choice').get('shuffle')
                    if randcheck == "yes":
                        file_html.write("<br><strong>Shuffle Qn? : {}   </strong><br>".format(randcheck))
                    elif randcheck == "no":
                        file_html.write("<br><strong>Shuffle Qn? : {}   </strong><br>".format(randcheck))
                    else:
                        file_html.write("<br><strong>shuffle unclear please check</strong></br>")

            file_html.write("</body><html>")
            file_html.close()

            if os.path.exists(directoryPath+"/"+NewFileName+".docx"):
                confirm = sg.popup_yes_no("A file with the name "+NewFileName+".docx already exists in "+directoryPath+". Do you want to replace it?")
                if confirm == "Yes":
                    
                    doc = aw.Document(tempFileDirec+"/"+NewFileName+".html")
                    doc.save(directoryPath+"/"+NewFileName+".docx")
                    shutil.rmtree(tempFileDirec)

                    sg.popup(".docx is saved at "+directoryPath+"/"+NewFileName+".docx")
                    os.startfile(directoryPath+"/"+NewFileName+".docx")
                else:
                    sg.popup("The new file was not saved.")
            else:
                
                doc = aw.Document(tempFileDirec+"/"+NewFileName+".html")
                doc.save(directoryPath+"/"+NewFileName+".docx")
                shutil.rmtree(tempFileDirec)

                sg.popup(".docx is saved at "+directoryPath+"/"+NewFileName+".docx")
                os.startfile(directoryPath+"/"+NewFileName+".docx")
            break


        except Exception as e:
            sg.popup("Something went wrong! Please check the exported file from NYPLMS! ")
            print(str(e))
            tb_str = traceback.format_tb(e.__traceback__)
            print(f"Error: {e} on line {tb_str[-1].split(',')[1]}")
            break


    if event ==sg.WIN_CLOSED or event=="Exit":
        break

    if event == 'User Manual':
        webbrowser.open('Z:/NYP_EQfromNYPLMSQuiz/User_Manual.pdf')
window.close()

    # WIP IMAGE FIX
    #         max_width = 4800
    #         ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    #               'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    #               'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

    #         for img in root.findall('.//w:drawing/w:pict/w:binData/../..', ns):
    # # get the original dimensions of the image
    #             cx = int(img.find('.//a:ext', ns).get('cx'))
    #             cy = int(img.find('.//a:ext', ns).get('cy'))
    #             print(cx,cy)
    # # modify the dimensions of the image
    #             cx /= 2
    #             cy /= 2

    # # update the XML file with the modified dimensions
    #             img.find('.//a:ext', ns).set('cx', str(cx))
    #             img.find('.//a:ext', ns).set('cy', str(cy))
class MultipleChoiceAndTF:
    def findscore(self):
        # for p in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}'):
            tempscore = p.find('setvar').text        
            try:
                formattedscore = int(float(tempscore))

            except:
                formattedscore = 0
            return formattedscore
    def findans(self,correct):                            
        # if type(formattedscore) is int and formattedscore > 0:
        #     for v in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}/conditionvar'):
        #         correct=v.find('varequal').text
                try:
                    AnsToBePrinted=ansIdList[correct]
                    if AnsToBePrinted.startswith("<p>"):
                        formattedAnsToBePrinted=AnsToBePrinted[3:-4]
                        file_html.write("<p>[ {} Marks ]</p>".format(int(float(qnmark_str))))
                        ansDict[qnNumber+1].append(formattedAnsToBePrinted)
                    else: 
                        ansDict[qnNumber+1].append(AnsToBePrinted)
                except:
                    pass
                return ansDict

class MultiSelect:
    def MultiSelQn(self,correct):
        # for p in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}/setvar'):
        #     score=p.attrib['varname']
        #     if score == "D2L_Correct":
        #         for v in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition{[loopCounter]}/conditionvar'):
        #             correct=v.find('varequal').text
                    try:
                        AnsToBePrinted=ansIdList[correct]
                    #print(AnsToBePrinted)
                        if AnsToBePrinted.startswith("<p>"):
                            formattedAnsToBePrinted=AnsToBePrinted[3:-4]
                        # file_html.write("<strong >Answer : {}  </strong><br>".format(formattedAnsToBePrinted))
                            ansDict[qnNumber+1].append(formattedAnsToBePrinted)
                        else:
                        # file_html.write("<strong>Answer : {}   </strong><br>".format(AnsToBePrinted))
                            ansDict[qnNumber+1].append(AnsToBePrinted)                    
                    except:
                        pass
                    return ansDict
class ShortAns:
    def ShortAns(self):
    # for o in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition/conditionvar'):
        try:
            LongAnswer=o.find('varequal').text
            file_html.write("<br><strong>Answer : {}   </strong><br>".format(LongAnswer))

        except:
            pass
        # else:
        #     break
class MultiShortAns:
    def MultiShortAns(self):
    # for o in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition/conditionvar'):
        try:
            LongAnswer=o.find('varequal').text
            file_html.write("<br><strong>Answer : {}   </strong><br>".format(LongAnswer))

        except:
            pass
class FTB:
    def FTB(self):
    # for o in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/resprocessing/respcondition/conditionvar'):
        try:
            LongAnswer=o.find('varequal').text
            file_html.write("<br><strong>Answer : {}   </strong><br>".format(LongAnswer))
        except:
            pass
class LongAns:
 def LongAns(self):
    temp=False
    for h in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/answer_key/answer_key_material/flow_mat/material'):
        LongAnswer=h.find('mattext').text
        temp=True
        if temp ==False:
                for h in root.findall(f'./objectbank/section/item[@ident="{QnList[qnNumber]}"]/answer_key/answer_key_material/flow_mat/flow_mat'):
                    LongAnswer=h.find('mattext').text
                    file_html.write("<p>[ {} Marks ]</p>".format(int(float(qnmark_str))))                                
        else:
            file_html.write("<p>[ {} Marks ]</p>".format(int(float(qnmark_str))))
            file_html.write("<br><strong>Answer : {}   </strong><br>".format(LongAnswer))
            break

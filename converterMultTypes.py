import streamlit as st
import pandas as pd
import os
import io
import csv
import zipfile
from io import BytesIO
from io import StringIO
import openpyxl
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Inches

class messages():
    def __init__(self, *args):
        self.fileTmp = args[0]
        self.suffix = args[1]
        self.nFiles = args[2]
        if self.nFiles == 1:
            self.fileFinal = f'zipado_sozinho_{self.suffix}.zip'
        else:
            self.fileFinal = f'zipado_múltiplos_{self.suffix}.zip'
        if None not in args:
            self.mensResult()
    
    def mensResult(self):
        colMens, colZip = st.columns([16, 5], width='stretch', vertical_alignment='center')
        colMens.success(f':blue[**{self.fileFinal}**] com ***{self.nFiles}*** arquivo(s). Clique no botão 👉.', 
                        icon='✔️')                              
        with open(self.fileTmp, "rb") as file:
            colZip.download_button(label='Download',
                                   data=file,
                                   file_name=self.fileFinal,
                                   mime='application/zip', 
                                   icon=':material/download:', 
                                   use_container_width=True)
    
    @st.dialog(' ')
    def configTwo(self, str):
        st.markdown(str)  
      
class downFiles():
    def __init__(self, *args):
        self.files = args[0]
        self.index = args[1]
        self.engine = args[2]
        self.ext = args[3]
        self.opt = args[4]
        self.filesZip = []
        self.nFiles = 0
        if self.index in [0, 1, 2]:
            if self.index in [0, 1]:
                self.csvXlsx() 
            elif self.index == 2:
                self.ext = 'txt'
                self.csvTsv()
        elif self.index == 3:
            self.csvDocx()     
        self.nameZip = f'arquivo_all_{self.ext}.zip'
        self.downZip()
        if os.path.getsize(self.nameZip) > 0:
            messages(self.nameZip, self.ext, self.nFiles)
            
    def csvXlsx(self):
        for f, file in enumerate(self.files):
            self.nameFile = file[0]
            self.dataFile = file[1]
            if self.opt == 2:
                self.fileOut = f'{self.nameFile}.xlsx'
            else:
                self.fileOut = f'{self.nameFile}.{self.ext}'
            self.sheetName = 'aba_única'
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title = 'dados'        
            for data in self.dataFile:
                try:
                    newData = [item.encode('ISO-8859-1').decode('utf8') for item in data]
                except: 
                    newData = [item for item in data]
                sheet.append(newData)
            wb.save(self.fileOut)
            self.bytesFiles(0)
    
    def csvTsv(self):
        for f, file in enumerate(self.files):
            self.nameFile = file[0]
            self.dataFile = file[1]
            self.fileOut = f'{self.nameFile}_new.csv'
            self.csvCsv()
            self.bytesFiles(1)
            
    def csvCsv(self):
        allLines = []
        for data in self.dataFile:
            try:
                newData = [item.encode('ISO-8859-1').decode('utf8') for item in data]
            except: 
                newData = [item for item in data]
            allLines.append(newData)
        with open(self.fileOut, 'w', newline='', encoding='utf-8') as recordCsv:
            writerCsv = csv.writer(recordCsv)
            writerCsv.writerows(allLines)
    
    def csvDocx(self):
        for f, file in enumerate(self.files):
            self.nameFile = file[0]
            self.dataFile = file[1]
            self.fileOut = f'{self.nameFile}_new.csv'
            self.csvCsv()
            doc = Document()
            with open(self.fileOut, 'r', encoding='utf-8') as f:
                csv_reader = csv.reader(f)
                headers = next(csv_reader)
                num_cols = len(headers)
                table = doc.add_table(rows=1, cols=num_cols)
                table.style = 'Table Grid'
                table.page_width = Inches(11.7) 
                table.page_height = Inches(8.5)                 
                hdr_cells = table.rows[0].cells
                for i, header in enumerate(headers):
                    hdr_cells[i].text = header
                for row in csv_reader:
                    row_cells = table.add_row().cells
                    for i, cell_data in enumerate(row):
                        row_cells[i].text = cell_data
                self.fileOut = f'{self.nameFile}_new.{self.ext}'
                doc.save(self.fileOut)
                self.bytesFiles(2)
            
    def bytesFiles(self, mode):
        if mode == 0:
            if self.opt != 2:
                self.df = pd.read_excel(self.fileOut).fillna('')
                self.renameHead()    
                output = BytesIO()
                with pd.ExcelWriter(output, engine=self.engine) as writer:
                    self.df.to_excel(writer, index=False, sheet_name=self.sheetName)
                excelBytes = output.getvalue()
                zips = (self.fileOut, excelBytes)
                self.filesZip.append(zips) 
                self.nFiles += 1
            else:
                self.df = pd.read_excel(self.fileOut).fillna('')
                self.renameHead()
                output = BytesIO()
                with pd.ExcelWriter(output, engine=self.engine) as writer:
                    self.df.to_excel(writer, index=False, sheet_name=self.sheetName)
                self.fileOut = f'{self.nameFile}.{self.ext}'              
                xlsxBytes = output.getvalue()
                bytesObj = io.BytesIO(xlsxBytes)
                self.df = pd.read_excel(bytesObj).fillna('')
                self.renameHead()
                htmlStr = self.df.to_html(index=False)
                htmlBytes = htmlStr.encode('utf-8')
                zips = (self.fileOut, htmlBytes)
                self.filesZip.append(zips) 
                self.nFiles += 1
        elif mode == 1:
            self.df = pd.read_csv(self.fileOut).fillna('')
            self.renameHead()
            output = BytesIO()
            self.df.to_csv(output, sep='\t', index=False)
            csvBytes = output.getvalue()
            self.fileOut = f'{self.nameFile}.txt'
            zips = (self.fileOut, csvBytes)
            self.filesZip.append(zips) 
            self.nFiles += 1
        elif mode == 2:
            newDocx = f'{self.nameFile}.{self.ext}'
            output = BytesIO()
            with open(self.fileOut, 'rb') as arquivo:
                docxRead = arquivo.read()
            zips = (newDocx, docxRead)
            self.filesZip.append(zips) 
            self.nFiles += 1
    
    def renameHead(self):
        head = {}
        for col in self.df.columns:
            if col.lower().find('unnamed') >= 0:
                head[col] = ''
            else:
                head[col] = col
        self.df.rename(columns=head, inplace=True)
        
    def downZip(self):
        with zipfile.ZipFile(self.nameZip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in self.filesZip:
                nameFile = file[0]
                dataFile = file[1]
                zipf.writestr(nameFile, dataFile)
        
class main():
    def __init__(self):          
        st.set_page_config(initial_sidebar_state="collapsed", layout="wide")
        self.typeExt = sorted(['CSV', 'XLS', 'XLSX'])   
        colType, colUpload = st.columns([12, 17], width='stretch')
        self.uploads = ['upLoad']
        with colType:            
            with st.container(border=4, key='contType', gap='small', height="stretch"):
                self.typeFile = st.selectbox('Selecione o tipo de arquivo', self.typeExt, key='typeFile') 
                self.upLoad = st.file_uploader(f'Escolha dois ou mais arquivos {self.typeFile}.', 
                                               type=self.typeFile, accept_multiple_files=True, key=self.uploads[0])    
        with colUpload:  
            try:
                self.files = list(set([file.name for file in self.upLoad]))
            except:
                self.files = [] 
            if not self.typeFile:
                with st.container(border=4, key='contZero', gap='small', height="stretch"):
                    st.text("")
            if self.typeFile:
                with st.container(border=None, key='contUpload', gap='small', height="stretch"):
                    nUploads = len(self.upLoad)
                    self.disableds = ['disabled' + str(w) for w in range(6)]
                    if nUploads == 0:
                        self.setSessionState(True, 0)
                    else:
                        self.setSessionState(False, 0)
                    if not self.typeFile:
                        with st.container(border=4, key='contOne', gap='small', height="stretch"):
                            st.text('')                            
                    else:
                        if not self.typeFile == self.typeExt[0]: 
                            with st.container(border=4, key='contOne', gap='small', height="stretch"):
                                st.text('')
                        else:
                            self.exts = {'openpyxl': ['xls', 'xlsx', 'html'], 
                                         'odf': ['ods'], 
                                         'tsv': ['tsv'], 
                                         'doc': ['docx']}                        
                            stripe = f':red[**{self.typeExt[1].lower()}**]'
                            with st.container(border=4, key='contOne', gap='small', height="stretch"):
                                nFiles = len(self.files)
                                if nFiles <= 1:
                                    titleSel = f'Arquivo sem duplicação ({nFiles})'
                                else:
                                    titleSel = f'Arquivos sem duplicação ({nFiles})'
                                self.fileSel = st.selectbox(titleSel, options=self.files, key='files')
                                colOne, colTwo, colThree = st.columns(spec=3, width='stretch')
                                buttOne = colOne.button(label=f'{stripe} para xls', key='buttOne',
                                                        use_container_width=True, 
                                                        icon=':material/sync_alt:', 
                                                        disabled=st.session_state[self.disableds[0]]) 
                                buttTwo = colTwo.button(label=f'{stripe} para xlsx', key='buttTwo',
                                                        use_container_width=True, 
                                                        icon=':material/swap_horiz:', 
                                                        disabled=st.session_state[self.disableds[1]])
                                buttThree = colThree.button(label=f'{stripe} para html', key='buttThree',
                                                        use_container_width=True, 
                                                        icon=':material/table_convert:', 
                                                        disabled=st.session_state[self.disableds[3]]) 
                                colFour, colFive, colSix = st.columns(spec=3, width='stretch')
                                buttFour = colFour.button(label=f'{stripe} para ods', key='buttFour',
                                                        use_container_width=True, 
                                                        icon=':material/transform:', 
                                                        disabled=st.session_state[self.disableds[3]]) 
                                buttFive = colFive.button(label=f'{stripe} para txt', key='buttFive',
                                                        use_container_width=True, 
                                                        icon=':material/convert_to_text:', 
                                                        disabled=st.session_state[self.disableds[4]])
                                buttSix = colSix.button(label=f'{stripe} para docx', key='buttSix',
                                                        use_container_width=True, 
                                                        icon=':material/convert_to_text:', 
                                                        disabled=st.session_state[self.disableds[5]])
                            if self.upLoad:  
                                if any([buttOne, buttTwo, buttThree, buttFour, buttFive, buttSix]):
                                    if buttOne:                     
                                        self.index = 0
                                        self.opt = 0
                                    elif buttTwo:
                                        self.index = 0
                                        self.opt = 1
                                    elif buttThree:
                                        self.index = 0
                                        self.opt = 2                        
                                    elif buttFour:
                                        self.index = 1
                                        self.opt = 0 
                                    elif buttFive: 
                                        self.index = 2
                                        self.opt = 0 
                                    elif buttSix: 
                                        self.index = 3
                                        self.opt = 0 
                                    self.keys = list(self.exts.keys())
                                    self.key = self.keys[self.index]
                                    self.values = self.exts[self.key]
                                    self.ext = self.values[self.opt] 
                                    self.filesRead = [] 
                                    self.segregateFiles()
                                    downFiles(self.filesRead, self.index, self.key, self.ext, self.opt)
                        
    def setSessionState(self, state, mode):
        for disabled in self.disableds:
            if disabled not in st.session_state:
                st.session_state[disabled] = True 
            else:
                st.session_state[disabled] = state
  
    def segregateFiles(self):
        st.write(self.files)
        filesFind = {}        
        for upLoad in self.upLoad: 
            nameGlobal = upLoad.name
            filesFind.setdefault(nameGlobal, 0)
            if nameGlobal in self.files:
                filesFind[nameGlobal] += 1
            if filesFind[nameGlobal] > 1:
                continue
            nameFile, ext = os.path.splitext(nameGlobal)
            newName = f'{nameFile}.{self.ext}'
            dataBytes = upLoad.getvalue()
            dataString = dataBytes.decode('ISO-8859-1')
            self.fileMemory = io.StringIO(dataString)
            sep = self.detectSep()
            readerCsv = csv.reader(self.fileMemory, delimiter=sep)
            joinNameRead = (nameFile, readerCsv)
            self.filesRead.append(joinNameRead)
            
    def detectSep(self):
        lines = 1024*10
        sample = self.fileMemory.read(lines)
        self.fileMemory.seek(0)
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
            
if __name__ == '__main__':
    with open('configCss.css) as f:
        css = f.read()
    st.markdown(f'<style>{css}</style>', unsafe_allow_html=True) 
    main()


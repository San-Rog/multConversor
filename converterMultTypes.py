import streamlit as st
import pandas as pd
import os
import io
import csv
import zipfile
from io import BytesIO
from io import StringIO
import openpyxl
import odf
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
        self.place = st.empty()
        self.place.write('')
        if self.nFiles <= 1:
            exprFile = ['arquivo', 'abri-lo']
        else:
            exprFile = ['arquivos', 'abri-los']
        if self.suffix in ['tsv']:
            mensStr = f':blue[**{self.fileFinal}**] contendo ***{self.nFiles}*** {exprFile[0]}. Clique no botão 👉.\n' \
                      f'(Use **Bloco de Notas** ou aplicativo similar para {exprFile[1]}.)'
        else:
            mensStr = f':blue[**{self.fileFinal}**] com ***{self.nFiles}*** {exprFile[0]}. Clique no botão 👉.'
        colMens, colZip = st.columns([17, 5], width='stretch', vertical_alignment='center')
        colMens.success(mensStr, icon='✔️')                              
        with open(self.fileTmp, "rb") as file:
            buttDown = colZip.download_button(label='Download',
                                   data=file,
                                   file_name=self.fileFinal,
                                   mime='application/zip', 
                                   icon=':material/download:', 
                                   use_container_width=True)
        if buttDown: 
            st.rerun()
    
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
                self.csvTsv()
        elif self.index == 3:
            self.csvDocx()     
        if self.opt is not None:
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
            self.csvCsv(0)
            self.bytesFiles(1)
               
    def csvPd(self):
        file = self.files[0]
        self.nameFile = file[0]
        self.dataFile = file[1]
        self.fileOut = self.nameFile
        self.csvCsv(1)
        self.df = pd.read_csv(self.fileOut).fillna('')
        self.renameHead()
        st.dataframe(self.df)
    
    def csvCsv(self, mode):
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
        if mode == 1:
            return self.fileOut
    
    def csvDocx(self):
        for f, file in enumerate(self.files):
            self.nameFile = file[0]
            self.dataFile = file[1]
            self.fileOut = f'{self.nameFile}_new.csv'
            self.csvCsv(0)
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
            self.fileOut = f'{self.nameFile}.{self.ext}'
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
        self.keyUp = 'zero'
        with colType:            
            with st.container(border=4, key='contType', gap='small', height="stretch"):
                self.typeFile = st.selectbox('Selecione o tipo de arquivo', self.typeExt, key='typeFile') 
                self.upLoad = st.file_uploader(f'Escolha dois ou mais arquivos {self.typeFile}.', 
                                               type=self.typeFile, accept_multiple_files=True, key=self.keyUp)    
        with colUpload:  
            try:
                self.files = list(set([file.name for file in self.upLoad]))
            except:
                self.files = [] 
            if not self.typeFile:
                self.configImageEmpty()
            if self.typeFile:
                with st.container(border=None, key='contUpload', gap='small', height="stretch"):
                    nUploads = len(self.upLoad)
                    self.disableds = ['disabled' + str(w) for w in range(6)]
                    if nUploads == 0:
                        self.setSessionState(True)
                    else:
                        self.setSessionState(False)
                    if not self.typeFile:
                        self.configImageEmpty()                           
                    else:
                        if not self.typeFile == self.typeExt[0]: 
                            self.configImageEmpty()
                        else:
                            self.exts = {'openpyxl': ['xls', 'xlsx', 'html'], 
                                         'odf': ['ods'], 
                                         'tsv': ['tsv'], 
                                         'doc': ['docx']}                        
                            self.newTypes = []
                            self.segregateTypes()
                            typeLow = self.typeFile.lower()
                            strFunc = ['Converter um ou mais arquivos', 'Convertendo']
                            stripe = f':red[**{self.typeFile.lower()}**]'
                            with st.container(border=4, key='contOne', gap='small', height="stretch"):
                                nFiles = len(self.files)
                                if nFiles <= 0:
                                    titleSel = f'Arquivo selecionado ({nFiles})'
                                else:
                                    titleSel = f'Arquivos selecionados ({nFiles})'
                                if nFiles > 0:
                                    opts = sorted(self.files)
                                    opts.insert(0, '')
                                else:
                                    opts = []
                                self.fileSel = st.selectbox(titleSel, options=opts, key='files', 
                                                            index=0)
                                colOne, colTwo, colThree = st.columns(spec=3, width='stretch')
                                buttOne = colOne.button(label=f'{stripe} para {self.newTypes[0]}', key='buttOne',
                                                        use_container_width=True, 
                                                        icon=':material/sync_alt:', 
                                                        disabled=st.session_state[self.disableds[0]],
                                                        help=f'{strFunc[0]} {stripe} para {self.newTypes[0]}.') 
                                buttTwo = colTwo.button(label=f'{stripe} para {self.newTypes[1]}', key='buttTwo',
                                                        use_container_width=True, 
                                                        icon=':material/swap_horiz:', 
                                                        disabled=st.session_state[self.disableds[1]], 
                                                        help=f'{strFunc[0]} {stripe} para {self.newTypes[1]}.')
                                buttThree = colThree.button(label=f'{stripe} para {self.newTypes[2]}', key='buttThree',
                                                        use_container_width=True, 
                                                        icon=':material/table_convert:', 
                                                        disabled=st.session_state[self.disableds[2]], 
                                                        help=f'{strFunc[0]} {stripe} para {self.newTypes[2]}.') 
                                colFour, colFive, colSix = st.columns(spec=3, width='stretch')
                                buttFour = colFour.button(label=f'{stripe} para {self.newTypes[3]}', key='buttFour',
                                                        use_container_width=True, 
                                                        icon=':material/transform:', 
                                                        disabled=st.session_state[self.disableds[3]], 
                                                        help=f'{strFunc[0]} {stripe} para {self.newTypes[3]}.') 
                                buttFive = colFive.button(label=f'{stripe} para {self.newTypes[4]}', key='buttFive',
                                                        use_container_width=True, 
                                                        icon=':material/convert_to_text:', 
                                                        disabled=st.session_state[self.disableds[4]], 
                                                        help=f'{strFunc[0]} {stripe} para {self.newTypes[4]}.')
                                buttSix = colSix.button(label=f'{stripe} para {self.newTypes[5]}', key='buttSix',
                                                        use_container_width=True, 
                                                        icon=':material/edit_arrow_up:',  
                                                        disabled=st.session_state[self.disableds[5]], 
                                                        help=f'{strFunc[0]} {stripe} para {self.newTypes[5]}.')
                                self.place = st.empty()
                                allButts = [buttOne, buttTwo, buttThree, buttFour, buttFive, buttSix]
                                if self.fileSel:
                                    self.place.write('')
                                    with st.spinner('Aguarde a exibição do arquivo na tela...'):
                                        nameFile = self.fileSel
                                        allNames = [file.name for file in self.upLoad]
                                        self.ext = self.typeFile.lower()
                                        self.pos = allNames.index(nameFile)
                                        self.filesReadDf = [] 
                                        self.segregateDf()
                                        objDown = downFiles(self.filesReadDf, None, None, self.ext, None)
                                        objDown.csvPd()    
                                if any(allButts):
                                    self.place.write('')
                            if self.upLoad:
                                if any(allButts):
                                    ind = allButts.index(True)
                                    expr = f'{strFunc[1]} {nUploads} do formato {stripe} para o foramto {self.newTypes[ind]}...'
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
                                    with st.spinner(expr):
                                        self.place.write('')
                                        self.keys = list(self.exts.keys())
                                        self.key = self.keys[self.index]
                                        self.values = self.exts[self.key]
                                        self.ext = self.values[self.opt] 
                                        self.filesRead = [] 
                                        self.segregateFiles()
                                        downFiles(self.filesRead, self.index, self.key, self.ext, self.opt)
                        
    def segregateTypes(self):
        listTypes = list(self.exts.values())
        for tipo in listTypes:
            self.newTypes += tipo
        self.newTypes = [f':red[**{new}**]' for new in self.newTypes]
    
    def configImageEmpty(self):
        with st.container(border=4, key='contZero', gap='small', height="stretch"):
            st.image(r'C:\Users\ACER\Downloads\image.jpg') 
    
    def setSessionState(self, state):
        for disabled in self.disableds:
            if disabled not in st.session_state:
                st.session_state[disabled] = True 
            else:
                st.session_state[disabled] = state
  
    def segregateFiles(self):
        filesFind = {}
        for upLoad in self.upLoad: 
            nameGlobal = upLoad.name
            nameFile, ext = os.path.splitext(nameGlobal)
            filesFind.setdefault(nameGlobal, 0)
            if nameGlobal in self.files:
                filesFind[nameGlobal] += 1
            if filesFind[nameGlobal] > 1:
                continue
            newName = f'{nameFile}.{self.ext}'
            dataBytes = upLoad.getvalue()
            dataString = dataBytes.decode('ISO-8859-1')
            self.fileMemory = io.StringIO(dataString)
            sep = self.detectSep()
            readerCsv = csv.reader(self.fileMemory, delimiter=sep)
            joinNameRead = (nameFile, readerCsv)
            self.filesRead.append(joinNameRead)
            
    def segregateDf(self):        
        for u, upLoad in enumerate(self.upLoad):
            if u == self.pos:
                nameGlobal = upLoad.name
                nameFile, ext = os.path.splitext(nameGlobal)
                newName = f'{nameFile}_new.{self.ext}'
                dataBytes = upLoad.getvalue()
                dataString = dataBytes.decode('ISO-8859-1')
                self.fileMemory = io.StringIO(dataString)
                sep = self.detectSep()
                readerCsv = csv.reader(self.fileMemory, delimiter=sep)
                joinNameRead = (nameFile, readerCsv)
                self.filesReadDf.append(joinNameRead)
                break
            
    def detectSep(self):
        lines = 1024*10
        sample = self.fileMemory.read(lines)
        self.fileMemory.seek(0)
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter

if __name__ == '__main__':
    with open('configCss.css') as f:
        css = f.read()
    st.markdown(f'<style>{css}</style>', unsafe_allow_html=True) 
    main()















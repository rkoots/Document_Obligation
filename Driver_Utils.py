## Author Rajkumar (C) rkoots@gmail.com # Rkoots 
import os,sys,re,time,pyPdf,xlsxwriter,subprocess
sys.path.append('docx2html')
from docx2html import convert
from HTMLParser import HTMLParser
from pyPdf.pdf import ContentStream
from pyPdf.pdf import TextStringObject

SLIDE_TEMPLATE = u'<p class="slide"><img src="{prefix}{src}" alt="{alt}" /></p>'
GHOSTSCRIPT = os.environ.get("GHOSTSCRIPT", "gs")
Keywords=["accountable","acting","actions","activities","adverse","agree","amendments","analyse","analysis","analyze","assign","assist","aware","cannot","categorize","classify","commence","communicate","compliance","comply","considers","contributing","coordinate","co-ordinating","create","decisions","declare","define","deliver","described","determine","develop","dispute","effects","efforts","ensure that","entitled  to","escalate","execute","follow","implement","issues","maintain","manage","minimize","monitor","must","objective","operate","oversee","perform","period","prepared","process","produce","propose","provide","provisions","raise","recommend","record","report","request","require","requires","reserves","resolution","resolve","resolved","respond to","responsible ","responsible for","result","review","set out","solely responsible","starting","submit","suggest","support","track","undertake","update" ]

lowriter = '/usr/bin/soffice'


def passdoc(file):
        File_data=[]
        class MyHTMLParser(HTMLParser):
                def handle_data(self, data):
                                File_data.append(data)
        parser = MyHTMLParser()
        parser.feed(file)
        return File_data

def CreateExcel(data,fileD):
        filename = 'DPE_'+str(time.strftime('%Y-%m-%d'))+'.xlsx'
        workbook = xlsxwriter.Workbook(filename)
        merge_format = workbook.add_format({
                'bold': 1,
                'num_format': 'YYYY-mm-dd',
                'border': 1,
                'align': 'center',
                'bg_color': '#003366',
                'color': 'white'})
        date_format = workbook.add_format({'num_format': 'YYYY-mm-dd','border': 1,'bold': 1, 'border': 1,'align': 'center','bg_color': '#003366','color': 'white'})
        date_format1 = workbook.add_format({'num_format': 'YYYY-mm-dd','align': 'center'})
        format1=workbook.add_format()
        format1.set_indent(1)
        format1.set_text_wrap()
        ECS=workbook.add_worksheet(filename.split(".")[0])
        ECS.set_column(1,15,20)
        header=["OBLIGATION_ID","DOCUMENT_REF_TYPE","MSA_SCHEDULE_NO","MSA_SOW_SECTION_NO","MSA_CLAUSE_NO","MSA_ANNEXURE_NO","MSA_APPENDICE_NO","MSA_APPENDIX_NO","MSA_ATTACHEMENT_NO","MSA_EXHIBIT_NO","MSA_PART_NO","MSA_REF","MSA_LEGAL_INTERPRETATION_BY_CLIENT_SP","OB_REVIEW_REMARKS","OB_DOMAIN_REF_ID"]
        col=1
        row=1
        for r in header:
                ECS.write(row,col,r,date_format)
                col=col+1
        row=row+1
        format1=workbook.add_format()
        format1.set_indent(1)
        i=1
        for ro in data:
                dummy_finder=0
                for df in range(0,7):
                        if ro[df].isdigit():
                                dummy_finder=df
                ro=ro[dummy_finder+1:]
                ECS.write(row,1,i,format1)
                ECS.write(row,2,fileD,format1)
                ECS.write(row,3,fileD,format1)
                ECS.write(row,4,"NA",format1)
                Claus=re.findall('\\d+', ro[:10])
                Clause=''
                for mat in Claus:
                        Clause=Clause+mat+'.'
                ECS.write(row,5,Clause,format1)
                ECS.write(row,6,"NA",format1)
                ECS.write(row,7,"NA",format1)
                ECS.write(row,8,"NA",format1)
                ECS.write(row,9,"NA",format1)
                ECS.write(row,10,"NA",format1)
                ECS.write(row,11,"1",format1)
                ECS.write(row,12,ro,format1)
                ECS.write(row,13,"",format1)
                ECS.write(row,14,"",format1)
                ECS.write(row,15,"",format1)
                row=row+1
                i+=1
        workbook.close()
        return filename

def extract_text(self):
    text = u""
    content = self["/Contents"].getObject()
    if not isinstance(content, ContentStream):
        content = ContentStream(content, self.pdf)
    for operands, operator in content.operations:
        if operator == "Tj":
            _text = operands[0]
            if isinstance(_text, TextStringObject):
                text += _text
        elif operator == "T*":
            text += "\n"
        elif operator == "'":
            text += "\n"
            _text = operands[0]
            if isinstance(_text, TextStringObject):
                text += operands[0]
        elif operator == '"':
            _text = operands[2]
            if isinstance(_text, TextStringObject):
                text += "\n"
                text += _text
        elif operator == "TJ":
            for i in operands[0]:
                if isinstance(i, TextStringObject):
                    text += i
        if text and not text.endswith(" "):
            text += " "  # Don't let words concatenate
    return text

def scrape_text(src):
    pages = []
    pdf = pyPdf.PdfFileReader(open(src, "rb"))
    for page in pdf.pages:
        text = extract_text(page)
        pages.append(text)
    return pages

def create_index_html(target, slides, prefix):
    out = open(target, "wt")
    print >> out, "<!doctype html>"
    for i in xrange(0, len(slides)):
        alt = slides[i]  # ALT text for this slide
        params = dict(src=u"slide%d.jpg" % (i+1), prefix=prefix, alt=alt)
        line = SLIDE_TEMPLATE.format(**params)
        print >> out, line.encode("utf-8")
    out.close()

def Patternmatch(Print):
        Set=[]
        for i in Print:
                if len(i.split())>4 and i[0].isdigit():
                        Flagger=0
                        for j in Keywords:
                                if j.lower() in i.lower():
                                        Flagger=1
                        if Flagger==1:
                                Set.append(i)
        return Set

def Method1(html,Var_E):
    Print = passdoc(html)
    result = Patternmatch(Print)
    CreateExcel(result,Var_E)

def Method2(html,Var_E):
    paragraphs = re.findall(r'(<p(.*?)</p>)', html)
    Print = []
    for i in paragraphs:
        Print.append(i[1][1:])
    result = Patternmatch(Print)
    CreateExcel(result, Var_E)

def DataComputer(html):
    Var_C = str(html)[:400]
    Var_D = re.findall(r'(<h2(.*?)</h2)', Var_C)
    Var_E = ''
    for allvar in Var_D:
        if ("CONFIDENTIAL" in allvar[0]) or ("CONTENTS" in allvar[0]):
            break
        if Var_E:
            Var_E = Var_E + " "
        Var_E = Var_E + allvar[1][1:]
    return Var_E

def DocxToPdf(abspath_pdf):
	subprocess.call('{0} --convert-to pdf "{1}" '.format(lowriter, abspath_pdf),shell=True)

def main():
    if len(sys.argv) < 3:
	 DocxToPdf(sys.argv[1])	
    if len(sys.argv) < 3:
        sys.exit("Usage: filename.py mypresentation.pdf / mypresentation.docx Method#")
    src = sys.argv[1]
    ParseMethod = sys.argv[2]
    if "docx" in src:
        html = convert(src)
        Var_E = DataComputer(html)
        if ParseMethod==1:
            Method1(html, Var_E)
        else:
            Method2(html, Var_E)
    else:
        basedir=os.path.dirname(os.path.realpath(__file__))
        pdfdir = os.path.normpath(basedir + '/pdf')
        docdir = os.path.normpath(basedir + '/doc')
        docxdir = os.path.normpath(basedir + '/docx')
        lowriter = '/usr/bin/soffice'
        outfilter = ':"MS Word 2007 XML"'
        outfilter = "'writer_pdf_import'"
        abspath_pdf = os.path.normpath(os.path.join(pdfdir,src))
        subprocess.call('{0} --infilter={1} --convert-to docx "{3}" --outdir "{2}"'.format(lowriter, outfilter, docxdir, abspath_pdf),shell=True)
        time.sleep(5)
        new_src=docxdir+'/'+src.split(".pdf")[0]+'.docx'
        html = convert(new_src)
        Var_E = DataComputer(html)
        if ParseMethod==1:
            Method1(html, Var_E)
        else:
            Method2(html, Var_E)
if __name__ == "__main__":
        main()

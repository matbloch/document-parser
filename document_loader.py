#!/usr/bin/env python3
import re
import subprocess
import os
import docxpy

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile


try:
    __import__('imp').find_module('eggs')
    PDF_SUPPORT = True
    import pdfparser.poppler as pdf
except ImportError:
    PDF_SUPPORT = False


class DocumentLoader:

    def read_document_content(self, filename):
        doc_type = self.get_type(filename)
        if doc_type == 'doc':
            return self.read_doc(filename)
        if doc_type == 'rtf':
            return self.read_rtf(filename)
        if doc_type == 'docx':
            return self.read_docx_content(filename, join_lines=True)
        if doc_type == 'pdf':
            return self.read_pdf(filename)
        raise NotImplementedError('Could not load file {} - implement reader for this document type'.format(filename))

    def get_type(self, filename):
        """
        Get the filetype - .doc/.pdf are not beeing validated
        TODO: filetype validation for all types (e.g. check binary signature)
        """

        base = os.path.basename(filename)
        ext = os.path.splitext(base)[1].lower()

        if ext == '.doc' or ext == '.rtf':
            if self.is_rtf(filename):
                return 'rtf'
            # DOC: NO validation!
            return 'doc'
        if ext == '.docx':
            if self.is_docx(filename):
                return 'docx'
        # PDF: NO validation!
        if ext == '.pdf':
            return 'pdf'
        raise NotImplementedError('Unknown filetype - extension: {}'.format(ext))

    # --- readers

    def read_doc(self, filename, method='antiword'):
        """
        method: antiword (prints lists but no headers), catdoc (prints headers but no lists)
        """
        if self.get_type(filename) != 'doc':
            raise ValueError('read_doc may only be used with .doc files.')

        if method == 'antiword':
            self.find_program('antiword', error_on_missing=True)
            method = 'antiword'
        elif method == 'catdoc':
            self.find_program('catdoc', error_on_missing=True)
            method = 'catdoc'
        else:
            raise NotImplementedError('Unsupported doc parsing method. Choose between "antiword" and "catdoc".')

        p = subprocess.Popen([method, filename], stdout=subprocess.PIPE)
        doc_text = p.stdout.read()
        return doc_text

    def read_rtf(self, filename):
        f = open(filename, "r")
        file_contents = f.read()
        return self.__striprtf(file_contents)

    def read_docx(self, filename):
        doc = docxpy.DOCReader(filename)
        doc.process()  # process file
        return doc.data

    def read_docx_content(self, filename, join_lines=True):
        texts = []
        for line in self.read_docx_rows(filename):
            texts.append(line)
        if join_lines:
            return '\n'.join(texts)
        else:
            return texts

    def read_pdf(self, filename, join_lines=True):
        if not PDF_SUPPORT:
            raise NotImplementedError('Please install "Poppler" to parse pdfs.')

        doc_reader = pdf.Document(filename, True, 0.0)  # @UndefinedVariable

        # print 'No of pages', doc_reader.no_of_pages
        # first page / last page
        # fp = args.first_page or 1
        # lp = args.last_page or d.no_of_pages

        content = []
        for p in doc_reader:
            # flow
            for f in p:
                # block
                for b in f:
                    # line
                    for l in b:
                        content.append(l.text.encode('UTF-8'))

        if join_lines:
            return '\n'.join(content)
        else:
            return content

    # --- helpers

    def is_rtf(self, filename):
        f = open(filename, "r")
        file_contents = f.read()
        if file_contents.startswith('{\\rtf'):
            return True
        return False

    def is_doc(self, filename):
        pass

    def is_docx(self, filename):
        f = zipfile.ZipFile(filename, "r")
        return self.isdir(f, "word") and self.isdir(f, "docProps") and self.isdir(f, "_rels")

    def isdir(self, z, name):
        return any(x.startswith("%s/" % name.rstrip("/")) for x in z.namelist())

    def read_docx_rows(self, path, keep_empty_lines=True):
        """
        Usage:
        texts = []
        for line in read_docx_rows(self.docx):
        texts.append(line)
        """

        WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        PARA = WORD_NAMESPACE + 'p'
        TEXT = WORD_NAMESPACE + 't'

        document = zipfile.ZipFile(path)
        xml_content = document.read('word/document.xml')
        document.close()
        tree = XML(xml_content)

        for paragraph in tree.getiterator(PARA):

            texts = [node.text
                     for node in paragraph.getiterator(TEXT)
                     if node.text]

            # print texts
            if keep_empty_lines:
                yield ''.join(texts).strip()
                continue

            if texts:
                text_preprocessed = ''.join(texts).strip()
                yield text_preprocessed
                continue

    def __striprtf(self, text):
        pattern = re.compile(r"\\([a-z]{1,32})(-?\d{1,10})?[ ]?|\\'([0-9a-f]{2})|\\([^a-z])|([{}])|[\r\n]+|(.)", re.I)
        # control words which specify a "destionation".
        destinations = frozenset((
            'aftncn', 'aftnsep', 'aftnsepc', 'annotation', 'atnauthor', 'atndate', 'atnicn', 'atnid',
            'atnparent', 'atnref', 'atntime', 'atrfend', 'atrfstart', 'author', 'background',
            'bkmkend', 'bkmkstart', 'blipuid', 'buptim', 'category', 'colorschememapping',
            'colortbl', 'comment', 'company', 'creatim', 'datafield', 'datastore', 'defchp', 'defpap',
            'do', 'doccomm', 'docvar', 'dptxbxtext', 'ebcend', 'ebcstart', 'factoidname', 'falt',
            'fchars', 'ffdeftext', 'ffentrymcr', 'ffexitmcr', 'ffformat', 'ffhelptext', 'ffl',
            'ffname', 'ffstattext', 'field', 'file', 'filetbl', 'fldinst', 'fldrslt', 'fldtype',
            'fname', 'fontemb', 'fontfile', 'fonttbl', 'footer', 'footerf', 'footerl', 'footerr',
            'footnote', 'formfield', 'ftncn', 'ftnsep', 'ftnsepc', 'g', 'generator', 'gridtbl',
            'header', 'headerf', 'headerl', 'headerr', 'hl', 'hlfr', 'hlinkbase', 'hlloc', 'hlsrc',
            'hsv', 'htmltag', 'info', 'keycode', 'keywords', 'latentstyles', 'lchars', 'levelnumbers',
            'leveltext', 'lfolevel', 'linkval', 'list', 'listlevel', 'listname', 'listoverride',
            'listoverridetable', 'listpicture', 'liststylename', 'listtable', 'listtext',
            'lsdlockedexcept', 'macc', 'maccPr', 'mailmerge', 'maln', 'malnScr', 'manager', 'margPr',
            'mbar', 'mbarPr', 'mbaseJc', 'mbegChr', 'mborderBox', 'mborderBoxPr', 'mbox', 'mboxPr',
            'mchr', 'mcount', 'mctrlPr', 'md', 'mdeg', 'mdegHide', 'mden', 'mdiff', 'mdPr', 'me',
            'mendChr', 'meqArr', 'meqArrPr', 'mf', 'mfName', 'mfPr', 'mfunc', 'mfuncPr', 'mgroupChr',
            'mgroupChrPr', 'mgrow', 'mhideBot', 'mhideLeft', 'mhideRight', 'mhideTop', 'mhtmltag',
            'mlim', 'mlimloc', 'mlimlow', 'mlimlowPr', 'mlimupp', 'mlimuppPr', 'mm', 'mmaddfieldname',
            'mmath', 'mmathPict', 'mmathPr', 'mmaxdist', 'mmc', 'mmcJc', 'mmconnectstr',
            'mmconnectstrdata', 'mmcPr', 'mmcs', 'mmdatasource', 'mmheadersource', 'mmmailsubject',
            'mmodso', 'mmodsofilter', 'mmodsofldmpdata', 'mmodsomappedname', 'mmodsoname',
            'mmodsorecipdata', 'mmodsosort', 'mmodsosrc', 'mmodsotable', 'mmodsoudl',
            'mmodsoudldata', 'mmodsouniquetag', 'mmPr', 'mmquery', 'mmr', 'mnary', 'mnaryPr',
            'mnoBreak', 'mnum', 'mobjDist', 'moMath', 'moMathPara', 'moMathParaPr', 'mopEmu',
            'mphant', 'mphantPr', 'mplcHide', 'mpos', 'mr', 'mrad', 'mradPr', 'mrPr', 'msepChr',
            'mshow', 'mshp', 'msPre', 'msPrePr', 'msSub', 'msSubPr', 'msSubSup', 'msSubSupPr', 'msSup',
            'msSupPr', 'mstrikeBLTR', 'mstrikeH', 'mstrikeTLBR', 'mstrikeV', 'msub', 'msubHide',
            'msup', 'msupHide', 'mtransp', 'mtype', 'mvertJc', 'mvfmf', 'mvfml', 'mvtof', 'mvtol',
            'mzeroAsc', 'mzeroDesc', 'mzeroWid', 'nesttableprops', 'nextfile', 'nonesttables',
            'objalias', 'objclass', 'objdata', 'object', 'objname', 'objsect', 'objtime', 'oldcprops',
            'oldpprops', 'oldsprops', 'oldtprops', 'oleclsid', 'operator', 'panose', 'password',
            'passwordhash', 'pgp', 'pgptbl', 'picprop', 'pict', 'pn', 'pnseclvl', 'pntext', 'pntxta',
            'pntxtb', 'printim', 'private', 'propname', 'protend', 'protstart', 'protusertbl', 'pxe',
            'result', 'revtbl', 'revtim', 'rsidtbl', 'rxe', 'shp', 'shpgrp', 'shpinst',
            'shppict', 'shprslt', 'shptxt', 'sn', 'sp', 'staticval', 'stylesheet', 'subject', 'sv',
            'svb', 'tc', 'template', 'themedata', 'title', 'txe', 'ud', 'upr', 'userprops',
            'wgrffmtfilter', 'windowcaption', 'writereservation', 'writereservhash', 'xe', 'xform',
            'xmlattrname', 'xmlattrvalue', 'xmlclose', 'xmlname', 'xmlnstbl',
            'xmlopen',
        ))
        # Translation of some special characters.
        specialchars = {
            'par': '\n',
            'sect': '\n\n',
            'page': '\n\n',
            'line': '\n',
            'tab': '\t',
            'emdash': '\u2014',
            'endash': '\u2013',
            'emspace': '\u2003',
            'enspace': '\u2002',
            'qmspace': '\u2005',
            'bullet': '\u2022',
            'lquote': '\u2018',
            'rquote': '\u2019',
            'ldblquote': '\201C',
            'rdblquote': '\u201D',
        }
        stack = []
        ignorable = False  # Whether this group (and all inside it) are "ignorable".
        ucskip = 1  # Number of ASCII characters to skip after a unicode character.
        curskip = 0  # Number of ASCII characters left to skip
        out = []  # Output buffer.
        for match in pattern.finditer(text.decode()):
            word, arg, hex, char, brace, tchar = match.groups()
            if brace:
                curskip = 0
                if brace == '{':
                    # Push state
                    stack.append((ucskip, ignorable))
                elif brace == '}':
                    # Pop state
                    ucskip, ignorable = stack.pop()
            elif char:  # \x (not a letter)
                curskip = 0
                if char == '~':
                    if not ignorable:
                        out.append('\xA0')
                elif char in '{}\\':
                    if not ignorable:
                        out.append(char)
                elif char == '*':
                    ignorable = True
            elif word:  # \foo
                curskip = 0
                if word in destinations:
                    ignorable = True
                elif ignorable:
                    pass
                elif word in specialchars:
                    out.append(specialchars[word])
                elif word == 'uc':
                    ucskip = int(arg)
                elif word == 'u':
                    c = int(arg)
                    if c < 0: c += 0x10000
                    if c > 127:
                        out.append(chr(c))  # NOQA
                    else:
                        out.append(chr(c))
                    curskip = ucskip
            elif hex:  # \'xx
                if curskip > 0:
                    curskip -= 1
                elif not ignorable:
                    c = int(hex, 16)
                    if c > 127:
                        out.append(chr(c))  # NOQA
                    else:
                        out.append(chr(c))
            elif tchar:
                if curskip > 0:
                    curskip -= 1
                elif not ignorable:
                    out.append(tchar)

        return ''.join(out)

    def find_program(self, prog_filename, error_on_missing=False):
        bdirs = ['$HOME/Environment/local/bin/',
                 '$HOME/bin/',
                 '/share/apps/bin/',
                 '/usr/local/bin/',
                 '/usr/bin/']
        paths_tried = []
        for d in bdirs:
            p = os.path.expandvars(os.path.join(d, prog_filename))
            paths_tried.append(p)
            if os.path.exists(p):
                return p
        if error_on_missing:
            raise Exception("*** ERROR: '%s' not found - please install it. Searched in:\n  %s\n" % (prog_filename, "\n  ".join(paths_tried)))
        else:
            return None
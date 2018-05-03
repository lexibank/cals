# coding=utf-8
"""
Note: We run libreoffice to convert from doc to docx after download.
"""
from __future__ import unicode_literals, print_function
from subprocess import check_call
import re
from collections import defaultdict

from docx import Document
from clldutils.dsv import UnicodeWriter, reader
from clldutils.misc import slug, lazyproperty
from clldutils.path import Path

from clldutils.path import Path
from pylexibank.dataset import Metadata
from pylexibank.dataset import Dataset as BaseDataset
from lingpy.sequence.sound_classes import clean_string


class Dataset(BaseDataset):
    dir = Path(__file__).parent

    def cmd_download(self, **kw):
        def rp(*names):
            return self.raw.posix(*names)

        fname = 'Table_S2_Supplementary_Mennecier_et_al..doc'
        self.raw.download_and_unpack(
            'https://ndownloader.figshare.com/articles/3443090/versions/1',
            fname,
            log=self.log)
        check_call(
            'libreoffice --headless --convert-to docx %s --outdir %s' % (rp(fname), rp()),
            shell=True)

        doc = Document(rp(Path(fname).stem + '.docx'))
        for i, table in enumerate(doc.tables):
            with UnicodeWriter(rp('%s.csv' % (i + 1,))) as writer:
                for row in table.rows:
                    writer.writerow(map(text_and_color, row.cells))

    def split_forms(self, row, value):
        return value.split(' ~ ')

    @lazyproperty
    def tokenizer(self):
        from segments import Tokenizer
        t = Tokenizer(profile=str(self.dir / 'etc' / 'orthography.tsv'))
        return lambda x, y: t(y).split()

    def cmd_install(self, **kw):
        gcode = {x['ID']: x['GLOTTOCODE'] for x in self.languages}
        ccode = {x.english: x.concepticon_id for x in self.conceptlist.concepts.values()}
        data = defaultdict(dict)
        for fname in self.raw.glob('*.csv'):
            read(fname, data)

        with self.cldf as ds:
            for doculect, wl in data.items():
                ds.add_language(
                    ID=slug(doculect), Name=doculect, Glottocode=gcode[doculect.split('-')[0]])

                for concept, (form, loan, cogset) in wl.items():
                    if concept in ccode:
                        csid = ccode[concept]
                    elif concept.startswith('to ') and concept[3:] in ccode:
                        csid = ccode[concept[3:]]
                    else:
                        csid = None
                    ds.add_concept(ID=slug(concept), Name=concept, Concepticon_ID=csid)
                    for row in ds.add_lexemes(
                            Language_ID=slug(doculect), Parameter_ID=slug(concept), Value=form):
                        if cogset:
                            ds.add_cognate(
                                lexeme=row,
                                Cognateset_ID='%s-%s' % (slug(concept), slug(cogset)))
                            break
            ds.align_cognates()


COLOR_PATTERN = re.compile('fill="(?P<color>[^"]+)"')


def text_and_color(cell):
    color = None
    for line in cell._tc.tcPr.xml.split('\n'):
        if 'w:shd' in line:
            m = COLOR_PATTERN.search(line)
            if m:
                color = m.group('color')
                break
    if color == 'auto':
        color = None
    if color:
        color = '#' + color + ' '
    return '%s%s' % (color if color else '', cell.paragraphs[0].text)


def get_loan_and_form(c):
    if c.startswith('#'):
        return c.split(' ', 1)
    return None, c


def read(fname, data):
    concepts, loan = None, None

    for i, row in enumerate(reader(fname)):
        if i == 0:
            concepts = {j: c for j, c in enumerate(row[1:])}
        else:
            for j, c in enumerate(row[1:]):
                if j % 2 == 0:  # even number
                    loan, form = get_loan_and_form(c)
                else:
                    if form.strip():
                        data[row[0]][concepts[j]] = (form, loan, c)
    return data
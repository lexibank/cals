# coding=utf-8
"""
Note: We run libreoffice to convert from doc to docx after download.
"""
from __future__ import unicode_literals, print_function
from subprocess import check_call
import re
import unicodedata
from collections import defaultdict

from docx import Document
from clldutils.dsv import UnicodeWriter, reader
from clldutils.misc import slug, lazyproperty
from clldutils.path import Path
from segments import Tokenizer, Profile

from clldutils.path import Path
from pylexibank.dataset import Metadata
from pylexibank.dataset import Dataset as BaseDataset
from lingpy.sequence.sound_classes import clean_string

SOURCE = 'Mennecier2016'


class Dataset(BaseDataset):
    dir = Path(__file__).parent
    id = 'cals'

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
        profile = self.dir / 'etc' / 'orthography.tsv'
        tokenizer = Tokenizer(profile=Profile.from_file(str(profile), form='NFC'))
        def _tokenizer(item, string, **kw):
            kw.setdefault("column", "Grapheme")
            kw.setdefault("separator", " _ ")
            return tokenizer(unicodedata.normalize('NFC', string), **kw).split()
        return _tokenizer

    def cmd_install(self, **kw):
        gcode = {x['ID']: x['Glottocode'] for x in self.languages}
        data = defaultdict(dict)
        for fname in self.raw.glob('*.csv'):
            read(fname, data)

        with self.cldf as ds:
            ds.add_sources()
            ccode = ds.add_concepts(id_factory=lambda c: slug(c.label))
            for doculect, wl in sorted(data.items()):
                sd = slug(doculect)
                
                ds.add_language(ID=sd, Name=doculect, Glottocode=gcode[doculect.split('-')[0]])
                for concept, (form, loan, cogset) in sorted(wl.items()):
                    sc = slug(concept)
                    if sc in ccode:
                        pass
                    elif sc.startswith('to ') and sc[3:] in ccode:
                        sc = sc[3:]
                    else:
                        sc = None
                    
                    for row in ds.add_lexemes(Language_ID=sd, Parameter_ID=sc, Value=form, Source=SOURCE):
                        if cogset:
                            ds.add_cognate(lexeme=row, Cognateset_ID='%s-%s' % (sc, slug(cogset)))
                            break
            #ds.align_cognates()


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

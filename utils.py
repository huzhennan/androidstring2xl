
import os
from re import sub, match
from android2po import program
from openpyxl.workbook import Workbook
from babel.messages import pofile
from xlutils import merge_xl_content, merge_component_xl
import xlutils

def find_file_dirs_iter(cur, filename):
    dirlist = []
    def callback(arg, dirname, names):
        for name in names:
            if name == filename:
                print name
                dirlist.append(dirname)

    os.path.walk(cur, callback, "")
    return dirlist.__iter__()

def filedir2name(filedir):
    return sub('/', '_', filedir)

def name2filedir(name):
    return sub('_', '/', name)

def language_po_file_type(code):
    return r'(.)*%s.po$' % code

def read_catalog(filename, **kwargs):
    """Helper to read a catalog from a .po file.
    """
    file = open(filename, 'rb')
    try:
        return pofile.read_po(file, **kwargs)
    finally:
        file.close()

def update_po_content2worksheet(filename, worksheet, local_code):
    #print 'filename: %s worksheet: %s' % (filename, worksheet)
    catalog = read_catalog(filename, locale=local_code)
    for msg in catalog:
        if not msg.id and not msg.context:
            continue

        # in xlsm file, A: source file name
        # B: -> Message 's msg.contenxt  -> xml's item's name
        # C: -> Message 's msg.id; msg.id is string -> xml's english's string.
        # D: -> Message msg.string.
        mA = unicode(os.path.basename(filename))
        mB = msg.context

        if isinstance(msg.id, unicode):
            mC = u"\"%s\"" % msg.id
        elif isinstance(msg.id, tuple) and len(msg.id) == 2:
            mC = u"(%s)" % "::".join(msg.id)
        else:
            mC = u""


        if isinstance(msg.string, unicode):
            mD = u"\"%s\"" % msg.string
        elif isinstance(msg.string, tuple):
            mD = u"(%s)" % "::".join(msg.string)
        else:
            mD = u""

        worksheet.append((mA, mB, mC, mD))


def visit_po_file(args, dirname, names):
    print 'args: %s \n dirname: %s \n names: %s' % \
          (args, dirname, names)

    ws = args['worksheet']
    language_code = args['language_code']
    filename_pattern = language_po_file_type(language_code)
    for name in names:
        subname = os.path.join(dirname, name)
        print 'match: %s ' % match(filename_pattern, name)
        if os.path.isfile(subname) and match(filename_pattern, name):
            update_po_content2worksheet(subname, ws, language_code)


COMPONET_XLSX_FILENAME = 'books.xlsx'

def write_po_content_to_xlsx(dir, language_code):
    wb = Workbook()
    dest_filename = os.path.join(dir, COMPONET_XLSX_FILENAME)

    ws = wb.worksheets[0]
    args = {'worksheet': ws, 'language_code':language_code}

    os.path.walk(dir, visit_po_file, args)

    wb.save(filename=dest_filename)


LANGUAGE_CODE = None
root_dir = os.getcwd()
local_dir = os.path.join(root_dir, "local")

def init_res():

    if not os.path.exists(local_dir):
        os.mkdir(local_dir)
    print 'local_dir: %s' % local_dir

    for componet_path in find_file_dirs_iter(os.path.abspath(root_dir), 'AndroidManifest.xml'):
        res_dir = os.path.join(componet_path, 'res')
        if not os.path.exists(res_dir):
            continue

        values_dir = os.path.join(res_dir, 'values')
        if not os.path.exists(values_dir):
            continue

        subdir_name = filedir2name(os.path.relpath(componet_path, root_dir))
        gettext_dir = os.path.join(local_dir, subdir_name)
        if not os.path.exists(gettext_dir):
            os.mkdir(gettext_dir)

        program.main(('test', 'init', '--android', res_dir, '--gettext', gettext_dir, 'language', LANGUAGE_CODE))

        write_po_content_to_xlsx(gettext_dir, LANGUAGE_CODE)

    merge_component_xl(local_dir, 'totol.xlsx')

TOTAL_XLSX_FILEPATH = os.path.join(os.getcwd(), 'totol.xlsx')

def import_res():
    subxlfilepaths = xlutils.separate_xl_content(TOTAL_XLSX_FILEPATH)
    print subxlfilepaths

    # delete all po file??
    # TODO: delete all po file??

    for subxlfile in subxlfilepaths:
        xlutils.upodate_component_xl_content2pofile(subxlfile)

        subxldir = os.path.dirname(subxlfile)
        relative_dir = os.path.relpath(subxldir, local_dir)
        componet_res_dir = os.path.join(root_dir,name2filedir(relative_dir), 'res')
        print "componet_dir: %s" % componet_res_dir

        program.main(('test', 'import', '--android', componet_res_dir,
        '--gettext', subxldir))

if __name__ == '__main__':
    import sys
    argv = sys.argv
    if not len(argv) == 3:
        print 'Use: name [init|import] [language]'
        sys.exit(0)

    action = argv[1]
    LANGUAGE_CODE = argv[2]
    if action == 'init':
        init_res()
    elif action == 'import':
        import_res()
    else:
        print 'Use: [name] [init|import] [language]'


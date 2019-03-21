import argparse
import os, sys

# add parent dir to sys path
curr_dir = os.path.dirname(os.path.join(os.getcwd(), __file__))
sys.path.append(os.path.normpath(os.path.join(curr_dir, '..')))

import document_loader


def main(file=None):
    d = document_loader.DocumentLoader()

    # check single file
    if file is not None:
        print "Document {} is of type '{}'".format(file, d.get_type(file))
        print "Content: \n"
        print d.read_document_content(file)
        sys.stdout.flush()
        return

    # check predefined test files
    files = ['docx.docx', 'rtf.doc', 'doc.doc', 'pdf.pdf']
    for f in files:
        print "============================="
        print "Document {} is of type '{}'".format(f, d.get_type(f))
        print "---------"
        print "File contents:\n---------"
        print d.read_document_content(f)
        sys.stdout.flush()


# ================================= #
#              Main

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--file')
    args = parser.parse_args()
    main(args.file)

import os
import ghostscript
import multiprocessing
import csv
from unidecode import unidecode


# Define path to dir containing PDF files to be converted to PDF/A
filepath = 'F:/pdf_a_test/'

def convert_pdf_to_pdfa(filepath, output):
    """Creates mirrored dir structure in dir arch/ and converts every PDF file to PDF/A 1-3"""
    for root, dirs, files in os.walk(filepath):
        for fname in files:
            if not fname.endswith('.pdf'):
                continue
            path = os.path.join(root, fname)
            relative_path = os.path.relpath(path, filepath)
            outpath = os.path.join(output, relative_path)
            outdir = os.path.dirname(outpath)
            if not os.path.exists(outdir):
                os.makedirs(outdir)
            fname_lower = fname.lower()
            norm_fname = (unidecode(fname_lower.replace(" ", "_")))
            # uncomment for PDF/A-1a, when on linux delete in command below win64c
            # ghostScriptExec = ['gswin64c', '-dPDFA=1a', '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE',
            #                    '-dCompatibilityLevel=1.4', '-dPDFACompatibilityPolicy=1',
            #                    '-sProcessColorModel=DeviceRGB', '-sColorConversionStrategy=UseDeviceIndependentColor',
            #                    '-sOutputICCProfile=esrgb.icc',
            #                    '-sDEVICE=pdfwrite',  '-sOutputFile=' + outpath+ '/arch_v_' + fname,
            #                    '-sPDFACompatibilityPolicy=1 "C:/Program Files/gs/lib/PDFA_def.ps"']

            # comment for if PDF/A-2b is unwanted, when on linux delete in command below win64c
            ghostScriptExec = ['gswin64c', '-dPDFA=2b', '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE',
                               '-dCompatibilityLevel=1.7', '-dPDFACompatibilityPolicy=1',
                               '-sProcessColorModel=DeviceRGB', '-sColorConversionStrategy=UseDeviceIndependentColor',
                               '-sOutputICCProfile=esrgb.icc', '-sDEVICE=pdfwrite',
                               '-sOutputFile=' + outdir + '/arch_v_' + norm_fname, path,
                               '-sPDFACompatibilityPolicy=1 "C:\Program Files\gs\lib\PDFA_def.ps"']

            # uncomment for PDF/A-3a, when on linux delete in command below win64c
            # ghostScriptExec = ['gswin64c', '-dPDFA=3a', '-dBATCH', '-dNOPAUSE', '-dNOOUTERSAVE',
            #                    '-dCompatibilityLevel=1.7', '-dPDFACompatibilityPolicy=1',
            #                    '-sProcessColorModel=DeviceRGB', '-sColorConversionStrategy=UseDeviceIndependentColor',
            #                    '-sOutputICCProfile=esrgb.icc', '-sDEVICE=pdfwrite',
            #                    '-sOutputFile=' + outpath + '/arch_v_' + fname, path,
            #                    '-sPDFACompatibilityPolicy=1 "C:\Program Files\gs\lib\PDFA_def.ps"']


            ghostscript.Ghostscript(*ghostScriptExec)


def write_pdf_list_to_csv(output, csv_file):
    """Writes all files inside /arch to csv file as simple report"""
    with open(csv_file, 'w', encoding='utf8') as f:
        writer = csv.writer(f, lineterminator='\n')
        writer.writerow(["Conversion with GPL Ghostscript 10.01.1 (2023-03-27)"])
        writer.writerow(
            ["Conversion program: PDF_to_PDF-A1-3.py (based on https://github.com/maarty1226/PDF_to_PDF-A1-3)"])
        writer.writerow([])

    for root, dirs, files in os.walk(output):
            for file in files:
                path = os.path.join(output, file)
                writer.writerow([path])

    f.close()


if __name__ == '__main__':
    print('\n PDF to PDF/A Conversion started...\n')

    output = (filepath + 'arch/')
    csv_file = (output + 'pdf_list.csv')
    p = multiprocessing.Process(target=convert_pdf_to_pdfa, args=(filepath, output))
    p.start()
    p.join()
    write_pdf_list_to_csv(output, csv_file)

    print('\n Conversion Completed \n')
import csv

from sge_job_accounting_entry import SGEJobAccountingEntry
from slurm_job_accounting_entry import SlurmJobAccountingEntry

class JobAccountingFile:

    # =====
    #
    # CLASSES
    #
    # =====
    class SlurmDialect(csv.Dialect):

        delimiter = '|'
        doublequote = False
        escapechar = '\\'
        lineterminator = '\n'
        quotechar = '"'
        quoting = csv.QUOTE_MINIMAL
        skipinitialspace = True
        strict = True

    csv.register_dialect("slurm", SlurmDialect)

    class SGEDialect(csv.Dialect):

        delimiter = ':'
        doublequote = False
        escapechar = '\\'
        lineterminator = '\n'
        quotechar = '"'
        quoting = csv.QUOTE_MINIMAL
        skipinitialspace = True
        strict = True

    csv.register_dialect("sge", SGEDialect)


    # File object of open file controlled by this object.
    fp = None

    # Fields of each line from a possible header.
    raw_line_fields = None

    def __init__(self, filename, dialect=None):

        self.fp = open(filename, "r")
        self.dialect = dialect

        # Do we need to autodetect the dialect?
        if self.dialect is None:
            self.dialect = self.get_dialect()

        # Determine the fields from the header, if any:
        if self.dialect == "sge":
            self.raw_line_fields = SGEJobAccountingEntry.SGE_ACCOUNTING_FIELDS
        elif self.dialect == "slurm":
            # Read first line to get fields for Slurm
            header_line = self.fp.readline()
            self.raw_line_fields = header_line.split('|')
        else:
            self.fp.close()
            raise ValueError


    def __iter__(self):
        if self.fp is None:
            raise StopIteration
        else:
            self.reader = csv.DictReader(self.fp, fieldnames=self.raw_line_fields, dialect=self.dialect)
            return self


    def next(self):
        line_dict = self.reader.next()

        if self.dialect == "sge":
            return SGEJobAccountingEntry(line_dict, self.dialect)
        elif self.dialect == "slurm":
            return SlurmJobAccountingEntry(line_dict, self.dialect)


    def __del__(self):
        if self.reader is not None:
            del self.reader
        if self.fp is not None:
            self.fp.close()



    def get_dialect(self):
        # Reads first line of the file and analyzes it to determine what job scheduler produced it.
        # It puts the line back after it reads it.

        # Read the first potentially header line.
        header_line = self.fp.readline()

        # Put the first line back.
        self.fp.seek(0)

        # Is it SGE?  There would be at least 44 colons in the string then.
        if header_line.count(':') >= 44:

            return "sge"

        # Is it Slurm?  There would be at least 5 pipes in it.
        elif header_line.count('|') >= 5:

            return "slurm"

        else:
            return None



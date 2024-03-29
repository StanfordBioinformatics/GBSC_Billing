import csv
import sys

from sge_job_accounting_entry import SGEJobAccountingEntry
from slurm_job_accounting_entry import SlurmJobAccountingEntry

class JobAccountingFile:

    # =====
    #
    # CLASSES
    #
    # =====
    class SlurmDialect_Pipe(csv.Dialect):

        delimiter = SlurmJobAccountingEntry.DELIMITER_PIPE
        doublequote = False
        escapechar = '\\'
        lineterminator = '\n'
        quotechar = '"'
        quoting = csv.QUOTE_MINIMAL
        skipinitialspace = True
        strict = True

    csv.register_dialect("slurm_pipe", SlurmDialect_Pipe)

    class SlurmDialect_Bang(csv.Dialect):

        delimiter = SlurmJobAccountingEntry.DELIMITER_BANG
        doublequote = False
        escapechar = '\\'
        lineterminator = '\n'
        quotechar = '"'
        quoting = csv.QUOTE_MINIMAL
        skipinitialspace = True
        strict = True

    csv.register_dialect("slurm_bang", SlurmDialect_Bang)

    class SlurmDialect_Hash(csv.Dialect):

        delimiter = SlurmJobAccountingEntry.DELIMITER_HASH
        doublequote = False
        escapechar = '\\'
        lineterminator = '\n'
        quotechar = '"'
        quoting = csv.QUOTE_MINIMAL
        skipinitialspace = True
        strict = True

    csv.register_dialect("slurm_hash", SlurmDialect_Hash)

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
        else:
            # Read first line to get fields for Slurm
            header_line = self.fp.readline().rstrip()

            if self.dialect == "slurm_pipe":
                self.raw_line_fields = header_line.split(SlurmJobAccountingEntry.DELIMITER_PIPE)
            elif self.dialect == "slurm_bang":
                self.raw_line_fields = header_line.split(SlurmJobAccountingEntry.DELIMITER_BANG)
            elif self.dialect == "slurm_hash":
                self.raw_line_fields = header_line.split(SlurmJobAccountingEntry.DELIMITER_HASH)
            else:
                print("Cannot determine dialect from file %s" % (filename), file=sys.stderr)
                self.fp.close()
                raise ValueError


    def __iter__(self):
        if self.fp is None:
            raise StopIteration
        else:
            self.reader = csv.DictReader(self.fp, fieldnames=self.raw_line_fields, dialect=self.dialect)
            return self


    def __next__(self):
        line_dict = next(self.reader)

        if self.dialect == "sge":
            return SGEJobAccountingEntry(line_dict, self.dialect)
        elif self.dialect == "slurm_pipe":
            return SlurmJobAccountingEntry(line_dict, self.dialect)
        elif self.dialect == "slurm_bang":
            return SlurmJobAccountingEntry(line_dict, self.dialect)
        elif self.dialect == "slurm_hash":
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

        # Is it Slurm?  There would be at least 5 of some delimiter in it.
        elif header_line.count(SlurmJobAccountingEntry.DELIMITER_PIPE) >= 5:
            return "slurm_pipe"

        elif header_line.count(SlurmJobAccountingEntry.DELIMITER_BANG) >= 5:
            return "slurm_bang"

        elif header_line.count(SlurmJobAccountingEntry.DELIMITER_HASH) >= 5:
            return "slurm_hash"

        else:
            return None



import calendar
import time

class JobAccountingEntry:

    @staticmethod
    def dict_get_int(dictionary, field):
        value = dictionary.get(field)

        if value is not None:
            return int(value)
        else:
            return None


    @staticmethod
    def dict_get_timestamp(dictionary, field):
        value = dictionary.get(field)

        if value is not None:
            return calendar.timegm(time.strptime(value,"%Y-%m-%dT%H:%M:%S"))
        else:
            return None


    def parse_line_dict(self, line_dict):
        pass


    def __init__(self, job_sched_line_dict, dialect):

        # Save the whole dictionary, just in case.
        self.raw_fields = job_sched_line_dict
        self.dialect = dialect

        # Fields to be read from the entry
        self.failed_code = None

        self.submission_time = None
        self.start_time = None
        self.end_time = None

        self.owner = None
        self.job_name = None
        self.account = None
        self.project = None
        self.node_list = None
        self.cpus = None
        self.wallclock = None
        self.job_id = None

        # Extract the above fields from the line dictionary.
        self.parse_line_dict(job_sched_line_dict)

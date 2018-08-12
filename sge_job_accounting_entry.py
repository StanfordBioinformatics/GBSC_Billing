
import job_accounting_entry

class SGEJobAccountingEntry(job_accounting_entry.JobAccountingEntry):

    # OGE accounting file column info:
    # http://manpages.ubuntu.com/manpages/lucid/man5/sge_accounting.5.html
    SGE_ACCOUNTING_FIELDS = (
        'qname', 'hostname', 'group', 'owner', 'job_name', 'job_number',  # Fields 0-5
        'account', 'priority', 'submission_time', 'start_time', 'end_time',  # Fields 6-10
        'failed', 'exit_status', 'ru_wallclock', 'ru_utime', 'ru_stime',  # Fields 11-15
        'ru_maxrss', 'ru_ixrss', 'ru_ismrss', 'ru_idrss', 'ru_isrss', 'ru_minflt', 'ru_majflt',  # Fields 16-22
        'ru_nswap', 'ru_inblock', 'ru_oublock', 'ru_msgsnd', 'ru_msgrcv', 'ru_nsignals',  # Fields 23-28
        'ru_nvcsw', 'ru_nivcsw', 'project', 'department', 'granted_pe', 'slots',  # Fields 29-34
        'task_number', 'cpu', 'mem', 'io', 'category', 'iow', 'pe_taskid', 'max_vmem', 'arid',  # Fields 35-43
        'ar_submission_time'  # Field 44
    )

    # OGE accounting failed codes which invalidate the accounting entry.
    # From https://arc.liv.ac.uk/SGE/htmlman/htmlman5/sge_status.html
    ACCOUNTING_FAILED_CODES = (1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 18, 19, 20, 21, 26, 27, 28, 29, 36, 38)

    def parse_line_dict(self, sge_line_dict):

        self.submission_time = self.dict_get_int(sge_line_dict, 'submission_time')

        # Fill in the object's fields from the data in the dict given.
        self.failed_code = int(sge_line_dict.get('failed'))

        job_failed = self.failed_code in self.ACCOUNTING_FAILED_CODES
        if job_failed:
            self.end_time = self.submission_time  # The only valid date in the record.
        else:
            self.end_time = self.dict_get_int(sge_line_dict, 'end_time')

        self.owner = sge_line_dict['owner']
        self.job_name = sge_line_dict['job_name']
        self.account = sge_line_dict['account']
        self.project = sge_line_dict['project']
        self.node_list = sge_line_dict['hostname']
        self.cpus = self.dict_get_int(sge_line_dict,'slots')
        self.wallclock = self.dict_get_int(sge_line_dict, 'ru_wallclock')
        self.job_id = self.dict_get_int(sge_line_dict, 'job_number')


import job_accounting_entry

class SlurmJobAccountingEntry(job_accounting_entry.JobAccountingEntry):

    def parse_line_dict(self, slurm_line_dict):

        # Fill in the object's fields from the data in the dict given.
        self.failed_code = 0  # TODO: get proper value for this failed code

        self.submission_time = self.dict_get_timestamp(slurm_line_dict, 'Submit')
        self.end_time = self.dict_get_timestamp(slurm_line_dict, 'End')

        self.owner = slurm_line_dict['User']
        self.job_name = slurm_line_dict['JobName']
        self.account = slurm_line_dict['Account']
        self.project = slurm_line_dict['WCKey']  # for future development
        self.node_list = slurm_line_dict['NodeList']
        # HACK FOR SEPT 2018: NCPUS was broken in this month by a misconfiguration in Slurm; ReqCPUS has the right data.
        self.cpus = self.dict_get_int(slurm_line_dict, 'ReqCPUS')
        self.wallclock = self.dict_get_int(slurm_line_dict, 'ElapsedRaw')
        self.job_id = self.dict_get_int(slurm_line_dict, 'JobIDRaw')




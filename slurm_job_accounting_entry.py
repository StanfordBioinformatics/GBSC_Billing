
import job_accounting_entry

class SlurmJobAccountingEntry(job_accounting_entry.JobAccountingEntry):

    # The delimiter to use when exporting Slurm accounting data using 'sacct'
    #  (The default for sacct, '|', doesn't work, because field Constraints can contain a pipe.)
    SLURMACCOUNTING_DELIMITER = '!'

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
        self.cpus = self.dict_get_int(slurm_line_dict, 'NCPUS')
        self.wallclock = self.dict_get_int(slurm_line_dict, 'ElapsedRaw')
        self.job_id = self.dict_get_int(slurm_line_dict, 'JobIDRaw')




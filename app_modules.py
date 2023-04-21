class ClinicianList:
    def __init__(self, app):
        self.app = app

    def query(self):
        self.table = list(
            self.app.db.select(
                '[Clinicians]', (
                    '[Entry]',
                    '[Prefix]',
                    '[First]',
                    '[Last]',
                    '[Designation]',
                    '[Phone]',
                    '[Fax]',
                    '[Email]',
                    '[Address 1]',
                    '[Address 2]',
                    '[City]',
                    '[State]',
                    '[Zip]',
                    '[Enrolled]',
                    '[Inactive]',
                    '[Comments]'
                ), None, -1
            )
        )
        self.names = []
        for entry in self.table:
            self.names.append(self.fname(entry['[Prefix]'], entry['[First]'], entry['[Last]'], entry['[Designation]']), '$l, $f')

    def fname(self, prefix: str, first: str, last: str, designation: str, pattern: str):
        #&p = prefix, &f = first, &l = last, &d = designation, other characters are returned as-is
        spanner = {'$p': prefix, '$f': first, '$l': last, '$d': designation}
        span = ''
        for i in range(0, len(pattern)):
            if pattern[i] != '$':
                span += pattern[i]
            else:
                span += spanner[pattern[i]+pattern[i+1]]
                i += 1
        return span
        #'Entry, Prefix, First, Last, Designation, Phone, Fax, Email, [Address 1], [Address 2], City, State, Zip, Enrolled, Inactive, Comments'

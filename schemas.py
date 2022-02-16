class Clinician:
  #def __init__(self, first, last, addr1, addr2, city, state, zip, phone, fax, email):
  def __init__(self, *args):
    if isinstance(args, dict):
      self.data = args
    else:
      self.data = {
        'first': args[0],
         'last': args[1],
        'addr1': args[2],
        'addr2': args[3],
         'city': args[4],
        'state': args[5],
          'zip': args[6],
        'phone': args[7],
          'fax': args[8],
        'email': args[9]
      }

class PatientSample:
  #def __init__(self, sampleID, first, last, clinician, cultureType, chartNumber, collectionDate, receiveDate, location, comments):
  def __init__(self, *args):
    if isinstance(args, dict):
      self.data = args
    else:
      self.data = {
              'sampleID': args[0],
                 'first': args[1],
                  'last': args[2],
             'clinician': args[3],
           'cultureType': args[4],
           'chartNumber': args[5],
        'collectionDate': args[6],
           'recieveDate': args[7],
              'location': args[8],
              'comments': args[9]
      }

class WaterlineSample:
  #def __init__(self, sampleID, clinician, shipDate, collectionDate, receiveDate, operatoryID, comments):
  def __init__(self, *args):
    if isinstance(args, dict):
      self.data = args
    else:
      self.data = {
              'sampleID': args[0],
             'clinician': args[1],
              'shipDate': args[2],
        'collectionDate': args[3],
           'recieveDate': args[4],
           'operatoryID': args[5],
              'comments': args[6]
      }
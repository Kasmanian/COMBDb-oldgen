class Clinician:
  def __init__(self, first, last, addr1, addr2, city, state, zip, phone, fax, email):
    self.first = first
    self.last = last
    self.addr1 = addr1
    self.addr2 = addr2
    self.city = city
    self.state = state
    self.zip = zip
    self.phone = phone
    self.fax = fax
    self.email = email

class PatientSample:
  def __init__(self, sampleID, first, last, clinician, cultureType, chartNumber, collectionDate, receiveDate, location, comments):
    self.sampleID = sampleID
    self.first = first
    self.last = last
    self.clinician = clinician
    self.cultureType = cultureType
    self.chartNumber = chartNumber
    self.collectionDate = collectionDate
    self.receiveDate = receiveDate
    self.location = location
    self.comments = comments

class WaterlineSample:
  def __init__(self, sampleID, clinician, shipDate, collectionDate, receiveDate, operatoryID, comments):
    self.sampleID = sampleID
    self.clinician = clinician
    self.shipDate = shipDate
    self.collectionDate = collectionDate
    self.receiveDate = receiveDate
    self.operatoryID = operatoryID
    self.comments = comments
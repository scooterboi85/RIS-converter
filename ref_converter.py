#this script converts a RIS file exported from an EndNote library into unformatted Word CWYW citations 
try:
    filename = input("Enter the name of the input RIS file exported from the EndNote library linked to the Word Document: ")
    content = ''
    with open(filename, 'r') as f:
        content = f.read()
except FileNotFoundError:
    print("The specified input file does not exist.")
except PermissionError:
    print("You don't have permissions to access the input file.")
except IOError:
    print("An error occurred while reading or writing the file.")

refs = content.split('ER  - ')
out = ''
print(refs)

for ref in refs:
    if (ref.strip() == ''):
        continue
    AU = ''
    PY = ''
    ID = ''
    #extract author name
    i1 = ref.find('AU  -')
    i2 = ref.find(', ', i1)
    if i1 != -1 and i2 != -1:
      AU = ref[i1+6:i2]
    #find publication year
    i1 = ref.find('PY  -')
    if (i1 != -1):
      PY = ref[i1+6:i1+10]
    #find ref number
    i1 = ref.find('ID  -')
    if (i1 != -1):
      ID = ref[i1+6:].strip()
    out += ('{' + AU + ', ' + PY + ' #' + ID + '}\n')
with open('output.txt', 'w') as file:
    file.write(out)
print("Done! Citations written to the file output.txt")


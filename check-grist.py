import urllib.request, base64, zipfile, io, sys

share_token = 'oPsT24RoZGssbMN'
auth = base64.b64encode(f'{share_token}:'.encode()).decode()
file_name = 'CertificateOfTrustForLivingTrust_Example.docx'
file_url = f'https://tools.kushkurriculum.org/nextcloud/public.php/dav/files/{share_token}/{file_name}'

print('Downloading...')
req = urllib.request.Request(file_url, headers={'Authorization': f'Basic {auth}'})
with urllib.request.urlopen(req) as r:
    original = r.read()
print(f'Downloaded {len(original)} bytes')

# Show what entries are present
with zipfile.ZipFile(io.BytesIO(original)) as zf:
    entries_before = zf.namelist()
print('Entries before:', entries_before)

# Rebuild ZIP without src.zip
output = io.BytesIO()
with zipfile.ZipFile(io.BytesIO(original)) as zin:
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'src.zip':
                print(f'  Removing: {item.filename}')
                continue
            zout.writestr(item, zin.read(item.filename))

cleaned = output.getvalue()
print(f'Cleaned size: {len(cleaned)} bytes')

# Verify
with zipfile.ZipFile(io.BytesIO(cleaned)) as zf:
    entries_after = zf.namelist()
print('Entries after:', entries_after)
assert 'src.zip' not in entries_after, 'src.zip still present!'

# Save locally instead since public share is read-only
local_path = r'd:\source\repos\PDFTemplateGenerator\CertificateOfTrustForLivingTrust_Example.docx'
with open(local_path, 'wb') as f:
    f.write(cleaned)
print(f'Saved clean file to: {local_path}')
print('Upload this file to Nextcloud to replace the existing one.')



from trilium_py.client import ETAPI

server_url='http://127.0.0.1:8080'
tk='<token>'
et=ETAPI(server_url,tk)

# clear evernote's tags
clear_tags = ['source','author','source_application','content_class','sourceUrl']
for tag in clear_tags:
    all_notes=et.search_note("#"+tag)
    notes=all_notes['results']
    for note in notes:
        for attr in note['attributes']:
            if attr['name'] in clear_tags:
                et.delete_attribute(attr['attributeId'])
                print(attr['attributeId']+' deleted.')

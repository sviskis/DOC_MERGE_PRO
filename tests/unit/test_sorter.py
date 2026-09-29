from docmerge.core.sorter import natural_key, sort_items
from docmerge.domain.enums import SortMode
from docmerge.domain.models import DocumentItem

def test_natural_sort(): assert sorted(['10.docx','2.docx','1.docx'],key=natural_key)==['1.docx','2.docx','10.docx']


def test_order_list_sort_is_absolute_and_does_not_renaturalize():
    items=[]
    for index,name in enumerate(('C.docx','A.docx','B.docx')):
        item=DocumentItem(f'C:/tmp/{name}'); item.filename=name; item.manual_order=index+1; items.append(item)
    ordered=sort_items(items,SortMode.ORDER_LIST)
    assert [x.filename for x in ordered]==['C.docx','A.docx','B.docx']
    assert [x.manual_order for x in ordered]==[1,2,3]


def test_manual_sort_uses_manual_order():
    items=[DocumentItem('C:/tmp/a.docx'),DocumentItem('C:/tmp/b.docx')]
    items[0].manual_order=2; items[1].manual_order=1
    assert [x.manual_order for x in sort_items(items,SortMode.MANUAL)]==[1,2]

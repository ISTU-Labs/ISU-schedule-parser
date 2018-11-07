from lxml import etree
import csv

INPUT_FILE = "./sched.xml"

dom = etree.parse(INPUT_FILE)

print(dom)

regexpNS = "http://exslt.org/regular-expressions"

NS = {'ex': "urn:schemas-microsoft-com:office:spreadsheet",
      "html": "http://www.w3.org/TR/REC-html40",
      "x2": "http://schemas.microsoft.com/office/excel/2003/xml",
      "o": "urn:schemas-microsoft-com:office:office",
      "x": "urn:schemas-microsoft-com:office:excel",
      're': regexpNS,
      "xsi": "http://www.w3.org/2001/XMLSchema-instance",
      "ss": "urn:schemas-microsoft-com:office:spreadsheet",
      'c': "urn:schemas-microsoft-com:office:component:spreadsheet"}


def xpath(query, name="result.txt", node=None):
    global dom, NS
    if node is None:
        node = dom
    result = node.xpath(query, namespaces=NS)
    try:
        out = open(name, "w")
        for row in result:
            out.write(etree.tostring(row, encoding=str, pretty_print=True))
            out.write("\n"+"-"*80+"\n")
        out.close()
    except:
        print("WARNING: cannot save result")
    print("Found {} rows for {}".format(len(result), name))
    return result


allrows = xpath('//ex:Row/ex:Cell/ex:Data/../..', 'shed.txt')
row = xpath('//ex:Row/ex:Cell/ex:Data[text()="Дни"]/../..', 'days.txt')
days = row[0]
print(days.tag)
directions = xpath(".//ex:Data/text()", node=days)
directions = [d for d in directions if d.startswith('Напр')]
# print(directions)
groups = xpath(
    './following-sibling::ex:Row[1]/ex:Cell/ex:Data/text()', node=days)
# print(groups)
grp = {key: val for key, val in zip(groups, directions)}
print(grp)
shrows = xpath('./following-sibling::ex:Row', node=days)
shrows = shrows[1:]
print(shrows[0])

lastrow = xpath(
    "./following-sibling::ex:Row/ex:Cell/ex:Data[re:test(., '^Директор', 'i')]/../..", node=days)[0]

rows = []
for row in shrows:
    if row != lastrow:
        rows.append(row)
    else:
        break
print(len(rows))

arr = []  # list of rows


class Row:
    def __init__(self, i, l, r):
        self.i = i
        self.l = l
        self.r = r

    def cells(self):
        return list(self.r)

    def __getitem__(self, index):
        return self.l[index]

    def interprete(self):
        xl = 0  # Index of target cell
        xc = 0  # Index of Excel cell
        cells = self.cells()

        #import pudb
        # pu.db

        while True:  # enumerate all target cells
            try:
                if self[xl] is not None:
                    xl += 1
                    continue
            except IndexError:
                pass
            try:
                cell = cells[xc]
            except IndexError:
                return
            try:
                data = cell[0].text
            except IndexError:
                data = ''
            # print(cell.attrib)
            numRows = cell.get(
                '{urn:schemas-microsoft-com:office:spreadsheet}MergeDown', 0)
            numRows = int(numRows)
            for j in range(numRows+1):
                arr[j+self.i][xl] = data.strip()
            xl += 1
            xc += 1
            while True:
                if not self.l:
                    break
                if self.l[-1]:
                    break
                self.l.pop(-1)

    def is_empty(self):
        for l in self.l:
            if l is not None and l != '':
                return False
        return True

    def __setitem__(self, index, value):
        return self.setCell(index, value)

    def setCell(self, index, data):
        l = self.l
        ll = len(l)
        dl = index-ll+1
        if dl > 0:
            [l.append(None) for j in range(dl)]
        if l[index] is None:
            l[index] = data
        else:
            raise ValueError("cell already has a value")

    def __str__(self):
        return "|".join(self.l)

    def equals(self, other):
        if other is None:
            return False

        for a, b in zip(self.l, other.l):
            if a is None:
                a = ''
            if b is None:
                b = ''
            if a != b:
                return False
        return True


def contract():
    new_arr = []
    prev = None
    for i, a in enumerate(arr):
        if a.equals(prev):
            continue
        if not a.is_empty():
            new_arr.append(a)
            prev = a
    return new_arr


for i, r in enumerate(rows):
    arr.append(Row(i, [], r))


for r in arr:
    r.interprete()

arr = contract()

# for r in arr:
#    print(r)
with open('sched.csv', 'w', newline='') as csvfile:
    spamwriter = csv.writer(csvfile,
                            quotechar="'",
                            delimiter=';',
                            quoting=csv.QUOTE_NONNUMERIC)
    for r in arr:
        spamwriter.writerow(r.l)

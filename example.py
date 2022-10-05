from oleshko_excel import Excel

ex = Excel('SomeTable', False)

a = ex.Column('a', 'Table a', width=40, horizontal='center')
c = ex.Column('c')
c.setHeading('heading c')

a.append(123)
a.append(34567)
a.append(54678)
a.append(4567890)

c.append('Some data c2')
c.append('asdasfrg')
c.append('dfhbty')
c.append('rtdythgdsrr')

ex.save()
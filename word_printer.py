from win32com import client


def print_item_header(word_app, current_row, current_item, prev_item, cables):
    doc = word_app.ActiveDocument
    table = doc.Tables(1)
    cable_type = current_item.get('cable_type','')
    cable = cables.get(cable_type, {})

    # Cable
    table.Cell(current_row, 1).Merge(table.Cell(current_row, 5))
    table.Cell(current_row, 1).Select()
    word_app.Selection.ParagraphFormat.Alignment = client.constants.wdAlignParagraphCenter
    word_app.Selection.Delete()
    table.Cell(current_row, 1).Range.Text = f"{current_item.get('cable_num','')}"
    

    table.Cell(current_row + 1, 1).Merge(table.Cell(current_row + 1, 5))
    table.Cell(current_row + 1, 1).Select()
    word_app.Selection.ParagraphFormat.Alignment = client.constants.wdAlignParagraphCenter
    word_app.Selection.Delete()
    table.Cell(current_row + 1, 1).Range.Text = f"{cable.get('name', '')}"

    # Items
    table.Cell(current_row + 2, 2).Range.Text = f"{prev_item.get('item_type','')} ({prev_item.get('item_id','')})"
    table.Cell(current_row + 2, 3).Range.Text = f"{current_item.get('item_type','')} ({current_item.get('item_id','')})"
    table.Cell(current_row + 2, 5).Range.Text = f"Шлейф {current_item.get('harness_num','')}"

    # References
    table.Cell(current_row + 3, 2).Range.Text = f"{prev_item.get('item_ref','')}"
    table.Cell(current_row + 3, 3).Range.Text = f"{current_item.get('item_ref','')}"


def print_IPK_IPR_IPD_IPP(word_app, current_row, current_item, prev_item, cables):
    doc = word_app.ActiveDocument
    table = doc.Tables(1)

    # Create rows
    table.Cell(current_row, 1).Select()
    word_app.Selection.InsertRowsBelow(8)
    word_app.Selection.Delete()

    # Print cable num and type and references
    print_item_header(word_app, current_row, current_item, prev_item, cables)

    cable_type = current_item.get('cable_type','')
    cable = cables.get(cable_type, {})
    cross_sec = cable.get('cross_sec', '')
    
    # Connections
    table.Cell(current_row + 4, 1).Range.Text = "1"
    table.Cell(current_row + 5, 1).Range.Text = "2"
    table.Cell(current_row + 6, 1).Range.Text = "Экран"

    table.Cell(current_row + 4, 2).Range.Text = "A1XT2:1"
    table.Cell(current_row + 5, 2).Range.Text = "A1XT2:2"
    table.Cell(current_row + 6, 2).Range.Text = "A1XT2:3"

    table.Cell(current_row + 4, 3).Range.Text = "A1XT1:1"
    table.Cell(current_row + 5, 3).Range.Text = "A1XT1:2"
    table.Cell(current_row + 6, 3).Range.Text = "A1XT1:3"

    table.Cell(current_row + 4, 4).Range.Text = f"{cross_sec} мм² {cable.get('core1', '')}"
    table.Cell(current_row + 5, 4).Range.Text = f"{cross_sec} мм² {cable.get('core2', '')}"
    table.Cell(current_row + 6, 4).Range.Text = "Экран"

    table.Cell(current_row + 4, 5).Range.Text = "+"
    table.Cell(current_row + 5, 5).Range.Text = "-"
    
    current_row += 8
    return current_row

def print_IPP_216(word_app, current_row, current_item, prev_item, cables):
    doc = word_app.ActiveDocument
    table = doc.Tables(1)

    # Create rows
    table.Cell(current_row, 1).Select()
    word_app.Selection.InsertRowsBelow(10)
    word_app.Selection.Delete()

    # Print cable num and type and references
    print_item_header(word_app, current_row, current_item, prev_item, cables)
    
    cable_type = current_item.get('cable_type','')
    cable = cables.get(cable_type, {})
    cross_sec = cable.get('cross_sec', '')

    # Connections
    prev_item_type = prev_item.get('item_type')
    if prev_item_type == 'ПУС':

        table.Cell(current_row + 4, 1).Range.Text = "1"
        table.Cell(current_row + 5, 1).Range.Text = "2"
        table.Cell(current_row + 6, 1).Range.Text = "Экран"

        table.Cell(current_row + 4, 2).Range.Text = "A1XT11:1"
        table.Cell(current_row + 5, 2).Range.Text = "A1XT12:1"
        table.Cell(current_row + 6, 2).Range.Text = "A1XT11:3"

        table.Cell(current_row + 4, 3).Range.Text = "XT7:1"
        table.Cell(current_row + 5, 3).Range.Text = "XT7:3"
        table.Cell(current_row + 6, 3).Range.Text = "XT7:5"

        table.Cell(current_row + 4, 4).Range.Text = f"{cross_sec} мм² {cable.get('core1', '')}"
        table.Cell(current_row + 5, 4).Range.Text = f"{cross_sec} мм² {cable.get('core2', '')}"
        table.Cell(current_row + 6, 4).Range.Text = "Экран"

        table.Cell(current_row + 4, 5).Split(1, 2)
        table.Cell(current_row + 5, 5).Split(1, 2)

        table.Cell(current_row + 4, 5).Range.Text = "A"
        table.Cell(current_row + 5, 5).Range.Text = "B"

        table.Cell(current_row + 4, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 4, 6).Merge(table.Cell(current_row + 5, 6))

    elif prev_item_type == 'ИПЭС':
        
        table.Cell(current_row + 4, 1).Range.Text = "1"
        table.Cell(current_row + 5, 1).Range.Text = "2"
        table.Cell(current_row + 6, 1).Range.Text = "3"
        table.Cell(current_row + 7, 1).Range.Text = "4"
        table.Cell(current_row + 8, 1).Range.Text = "Экран"

        table.Cell(current_row + 4, 2).Range.Text = "X4:1"
        table.Cell(current_row + 5, 2).Range.Text = "X4:2"
        table.Cell(current_row + 6, 2).Range.Text = "X4:4"
        table.Cell(current_row + 7, 2).Range.Text = "X4:5"
        table.Cell(current_row + 8, 2).Range.Text = "X4:Корпус"

        table.Cell(current_row + 4, 3).Range.Text = "XT8:1"
        table.Cell(current_row + 5, 3).Range.Text = "XT8:3"
        table.Cell(current_row + 6, 3).Range.Text = "XT7:1"
        table.Cell(current_row + 7, 3).Range.Text = "XT7:3"
        table.Cell(current_row + 8, 3).Range.Text = "XT7:5"

        table.Cell(current_row + 4, 4).Range.Text = f"{cross_sec} мм² {cable.get('core1', '')}"
        table.Cell(current_row + 5, 4).Range.Text = f"{cross_sec} мм² {cable.get('core2', '')}"
        table.Cell(current_row + 6, 4).Range.Text = f"{cross_sec} мм² {cable.get('core3', '')}"
        table.Cell(current_row + 7, 4).Range.Text = f"{cross_sec} мм² {cable.get('core4', '')}"
        table.Cell(current_row + 8, 4).Range.Text = "Экран"

        table.Cell(current_row + 4, 5).Split(1, 2)
        table.Cell(current_row + 5, 5).Split(1, 2)
        table.Cell(current_row + 6, 5).Split(1, 2)
        table.Cell(current_row + 7, 5).Split(1, 2)

        table.Cell(current_row + 4, 5).Range.Text = "пит.+"
        table.Cell(current_row + 5, 5).Range.Text = "пит.-"
        table.Cell(current_row + 6, 5).Range.Text = "A"
        table.Cell(current_row + 7, 5).Range.Text = "B"

        table.Cell(current_row + 4, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 6, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 4, 6).Merge(table.Cell(current_row + 5, 6))
        table.Cell(current_row + 6, 6).Merge(table.Cell(current_row + 7, 6))
    
    else:
        table.Cell(current_row + 4, 1).Range.Text = "1"
        table.Cell(current_row + 5, 1).Range.Text = "2"
        table.Cell(current_row + 6, 1).Range.Text = "3"
        table.Cell(current_row + 7, 1).Range.Text = "4"
        table.Cell(current_row + 8, 1).Range.Text = "Экран"

        table.Cell(current_row + 4, 2).Range.Text = "XT8:2"
        table.Cell(current_row + 5, 2).Range.Text = "XT8:4"
        table.Cell(current_row + 6, 2).Range.Text = "XT7:2"
        table.Cell(current_row + 7, 2).Range.Text = "XT7:4"
        table.Cell(current_row + 8, 2).Range.Text = "XT7:6"

        table.Cell(current_row + 4, 3).Range.Text = "XT8:1"
        table.Cell(current_row + 5, 3).Range.Text = "XT8:3"
        table.Cell(current_row + 6, 3).Range.Text = "XT7:1"
        table.Cell(current_row + 7, 3).Range.Text = "XT7:3"
        table.Cell(current_row + 8, 3).Range.Text = "XT7:5"

        table.Cell(current_row + 4, 4).Range.Text = f"{cross_sec} мм² {cable.get('core1', '')}"
        table.Cell(current_row + 5, 4).Range.Text = f"{cross_sec} мм² {cable.get('core2', '')}"
        table.Cell(current_row + 6, 4).Range.Text = f"{cross_sec} мм² {cable.get('core3', '')}"
        table.Cell(current_row + 7, 4).Range.Text = f"{cross_sec} мм² {cable.get('core4', '')}"
        table.Cell(current_row + 8, 4).Range.Text = "Экран"

        table.Cell(current_row + 4, 5).Split(1, 2)
        table.Cell(current_row + 5, 5).Split(1, 2)
        table.Cell(current_row + 6, 5).Split(1, 2)
        table.Cell(current_row + 7, 5).Split(1, 2)

        table.Cell(current_row + 4, 5).Range.Text = "пит.+"
        table.Cell(current_row + 5, 5).Range.Text = "пит.-"
        table.Cell(current_row + 6, 5).Range.Text = "A"
        table.Cell(current_row + 7, 5).Range.Text = "B"

        table.Cell(current_row + 4, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 6, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 4, 6).Merge(table.Cell(current_row + 5, 6))
        table.Cell(current_row + 6, 6).Merge(table.Cell(current_row + 7, 6))

    current_row += 10
    return current_row


def print_IPES(word_app, current_row, current_item, prev_item, cables):
    doc = word_app.ActiveDocument
    table = doc.Tables(1)

    # Create rows
    table.Cell(current_row, 1).Select()
    word_app.Selection.InsertRowsBelow(10)
    word_app.Selection.Delete()

    # Print cable num and type and references
    print_item_header(word_app, current_row, current_item, prev_item, cables)
    
    cable_type = current_item.get('cable_type','')
    cable = cables.get(cable_type, {})
    cross_sec = cable.get('cross_sec', '')

    # Connections
    prev_item_type = prev_item.get('item_type')
    if prev_item_type == 'ПУС':

        table.Cell(current_row + 4, 1).Range.Text = "1"
        table.Cell(current_row + 5, 1).Range.Text = "2"
        table.Cell(current_row + 6, 1).Range.Text = "Экран"

        table.Cell(current_row + 4, 2).Range.Text = "A1XT11:1"
        table.Cell(current_row + 5, 2).Range.Text = "A1XT12:1"
        table.Cell(current_row + 6, 2).Range.Text = "A1XT11:3"

        table.Cell(current_row + 4, 3).Range.Text = "X3:4"
        table.Cell(current_row + 5, 3).Range.Text = "X3:5"
        table.Cell(current_row + 6, 3).Range.Text = "X3:Корпус"

        table.Cell(current_row + 4, 4).Range.Text = f"{cross_sec} мм² {cable.get('core1', '')}"
        table.Cell(current_row + 5, 4).Range.Text = f"{cross_sec} мм² {cable.get('core2', '')}"
        table.Cell(current_row + 6, 4).Range.Text = "Экран"

        table.Cell(current_row + 4, 5).Split(1, 2)
        table.Cell(current_row + 5, 5).Split(1, 2)

        table.Cell(current_row + 4, 5).Range.Text = "A"
        table.Cell(current_row + 5, 5).Range.Text = "B"

        table.Cell(current_row + 4, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 4, 6).Merge(table.Cell(current_row + 5, 6))

    elif prev_item_type == 'ИП 216-001Ех':
        
        table.Cell(current_row + 4, 1).Range.Text = "1"
        table.Cell(current_row + 5, 1).Range.Text = "2"
        table.Cell(current_row + 6, 1).Range.Text = "3"
        table.Cell(current_row + 7, 1).Range.Text = "4"
        table.Cell(current_row + 8, 1).Range.Text = "Экран"

        table.Cell(current_row + 4, 2).Range.Text = "XT8:2"
        table.Cell(current_row + 5, 2).Range.Text = "XT8:4"
        table.Cell(current_row + 6, 2).Range.Text = "XT7:2"
        table.Cell(current_row + 7, 2).Range.Text = "XT7:4"
        table.Cell(current_row + 8, 2).Range.Text = "XT7:6"

        table.Cell(current_row + 4, 3).Range.Text = "X3:1"
        table.Cell(current_row + 5, 3).Range.Text = "X3:2"
        table.Cell(current_row + 6, 3).Range.Text = "X3:4"
        table.Cell(current_row + 7, 3).Range.Text = "X3:5"
        table.Cell(current_row + 8, 3).Range.Text = "X3:Корпус"

        table.Cell(current_row + 4, 4).Range.Text = f"{cross_sec} мм² {cable.get('core1', '')}"
        table.Cell(current_row + 5, 4).Range.Text = f"{cross_sec} мм² {cable.get('core2', '')}"
        table.Cell(current_row + 6, 4).Range.Text = f"{cross_sec} мм² {cable.get('core3', '')}"
        table.Cell(current_row + 7, 4).Range.Text = f"{cross_sec} мм² {cable.get('core4', '')}"
        table.Cell(current_row + 8, 4).Range.Text = "Экран"

        table.Cell(current_row + 4, 5).Split(1, 2)
        table.Cell(current_row + 5, 5).Split(1, 2)
        table.Cell(current_row + 6, 5).Split(1, 2)
        table.Cell(current_row + 7, 5).Split(1, 2)

        table.Cell(current_row + 4, 5).Range.Text = "пит.+"
        table.Cell(current_row + 5, 5).Range.Text = "пит.-"
        table.Cell(current_row + 6, 5).Range.Text = "A"
        table.Cell(current_row + 7, 5).Range.Text = "B"

        table.Cell(current_row + 4, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 6, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 4, 6).Merge(table.Cell(current_row + 5, 6))
        table.Cell(current_row + 6, 6).Merge(table.Cell(current_row + 7, 6))
    
    else:
        table.Cell(current_row + 4, 1).Range.Text = "1"
        table.Cell(current_row + 5, 1).Range.Text = "2"
        table.Cell(current_row + 6, 1).Range.Text = "3"
        table.Cell(current_row + 7, 1).Range.Text = "4"
        table.Cell(current_row + 8, 1).Range.Text = "Экран"

        table.Cell(current_row + 4, 2).Range.Text = "X1:2"
        table.Cell(current_row + 5, 2).Range.Text = "X1:4"
        table.Cell(current_row + 6, 2).Range.Text = "X1:5"
        table.Cell(current_row + 7, 2).Range.Text = "X1:6"
        table.Cell(current_row + 8, 2).Range.Text = "X1:7"

        table.Cell(current_row + 4, 3).Range.Text = "X3:1"
        table.Cell(current_row + 5, 3).Range.Text = "X3:2"
        table.Cell(current_row + 6, 3).Range.Text = "X3:4"
        table.Cell(current_row + 7, 3).Range.Text = "X3:5"
        table.Cell(current_row + 8, 3).Range.Text = "X3:Корпус"

        table.Cell(current_row + 4, 4).Range.Text = f"{cross_sec} мм² {cable.get('core1', '')}"
        table.Cell(current_row + 5, 4).Range.Text = f"{cross_sec} мм² {cable.get('core2', '')}"
        table.Cell(current_row + 6, 4).Range.Text = f"{cross_sec} мм² {cable.get('core3', '')}"
        table.Cell(current_row + 7, 4).Range.Text = f"{cross_sec} мм² {cable.get('core4', '')}"
        table.Cell(current_row + 8, 4).Range.Text = "Экран"

        table.Cell(current_row + 4, 5).Split(1, 2)
        table.Cell(current_row + 5, 5).Split(1, 2)
        table.Cell(current_row + 6, 5).Split(1, 2)
        table.Cell(current_row + 7, 5).Split(1, 2)

        table.Cell(current_row + 4, 5).Range.Text = "пит.+"
        table.Cell(current_row + 5, 5).Range.Text = "пит.-"
        table.Cell(current_row + 6, 5).Range.Text = "A"
        table.Cell(current_row + 7, 5).Range.Text = "B"

        table.Cell(current_row + 4, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 6, 6).Range.Text = "Витая пара"
        table.Cell(current_row + 4, 6).Merge(table.Cell(current_row + 5, 6))
        table.Cell(current_row + 6, 6).Merge(table.Cell(current_row + 7, 6))

    current_row += 10
    return current_row


def print_end_item(word_app, current_row):
    doc = word_app.ActiveDocument
    table = doc.Tables(1)

    # Create rows
    table.Cell(current_row, 1).Select()
    word_app.Selection.InsertRowsBelow(6)
    word_app.Selection.Delete()

    # Cable
    table.Cell(current_row, 1).Merge(table.Cell(current_row, 5))
    table.Cell(current_row, 1).Select()
    word_app.Selection.ParagraphFormat.Alignment = client.constants.wdAlignParagraphCenter
    word_app.Selection.Delete()
    table.Cell(current_row, 1).Range.Text = "ХХХХХХ"

    table.Cell(current_row + 1, 1).Merge(table.Cell(current_row + 1, 5))
    table.Cell(current_row + 1, 1).Select()
    word_app.Selection.ParagraphFormat.Alignment = client.constants.wdAlignParagraphCenter
    word_app.Selection.Delete()
    table.Cell(current_row + 1, 1).Range.Text = "Последний кабель до ПУС"

    # Items
    table.Cell(current_row + 2, 2).Range.Text = "ХХХ"
    table.Cell(current_row + 2, 3).Range.Text = "ПУС (xxx)"
    table.Cell(current_row + 2, 5).Range.Text = "Шлейф ?"

    current_row += 6
    return current_row
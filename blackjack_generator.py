#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор файла Blackjack_Formulas_Only.xlsx
Игра Блэкджек в Excel только на формулах, без VBA
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.workbook.defined_name import DefinedName

def add_defined_name(wb, name, formula, local_sheet=None):
    """Добавляет именованный диапазон с формулой"""
    defn = DefinedName(name, attr_text=formula)
    if local_sheet:
        defn.localSheetId = local_sheet
    wb.defined_names[name] = defn

def create_blackjack_file(filename='Blackjack_Formulas_Only.xlsx'):
    # Создаем новую книгу
    wb = Workbook()
    ws = wb.active
    ws.title = 'Game'

    # Настройка стилей
    title_font = Font(bold=True, size=14)
    cell_alignment = Alignment(horizontal='left', vertical='center')

    # === ЗАГОЛОВКИ И ВХОДНЫЕ ЯЧЕЙКИ ===
    ws['A1'] = 'Seed для новой игры'
    ws['B1'] = 12345
    ws['A2'] = 'Карт игрока'
    ws['B2'] = 2
    ws['A3'] = 'Stand (TRUE/FALSE)'
    ws['B3'] = False

    # Форматирование входных ячеек
    for cell in ['B1', 'B2', 'B3']:
        ws[cell].font = Font(size=11, bold=True)
        ws[cell].alignment = cell_alignment

    # Проверка данных для Seed
    from openpyxl.worksheet.datavalidation import DataValidation
    dv_seed = DataValidation(type='whole', operator='between', formula1='1', formula2='1000000', allow_blank=False)
    dv_seed.error = 'Введите целое число от 1 до 1000000'
    dv_seed.errorTitle = 'Неверное значение Seed'
    ws.add_data_validation(dv_seed)
    dv_seed.add('B1')

    # Проверка данных для карт игрока
    dv_cards = DataValidation(type='whole', operator='between', formula1='2', formula2='20', allow_blank=False)
    dv_cards.error = 'Введите целое число от 2 до 20'
    dv_cards.errorTitle = 'Неверное количество карт'
    ws.add_data_validation(dv_cards)
    dv_cards.add('B2')

    # Проверка данных для Stand
    dv_stand = DataValidation(type='list', formula1='"TRUE,FALSE"', allow_blank=False)
    dv_stand.error = 'Введите TRUE или FALSE'
    dv_stand.errorTitle = 'Неверное значение Stand'
    ws.add_data_validation(dv_stand)
    dv_stand.add('B3')

    # Разделитель
    ws['A5'] = '=' * 50

    # === БЛОК РЕЗУЛЬТАТОВ ===
    ws['A7'] = '=== РЕЗУЛЬТАТЫ ИГРЫ ==='
    ws['A7'].font = title_font

    ws['A9'] = 'Рука игрока:'
    ws['C9'] = '=PlayerHandDisplay'

    ws['A10'] = 'Очки игрока:'
    ws['C10'] = '=PlayerTotal'

    ws['A12'] = 'Карта дилера (видимая):'
    ws['C12'] = '=DealerVisibleDisplay'

    ws['A13'] = 'Рука дилера (полная):'
    ws['C13'] = '=DealerHandDisplay'

    ws['A14'] = 'Очки дилера:'
    ws['C14'] = '=DealerTotal'

    ws['A16'] = 'РЕЗУЛЬТАТ:'
    ws['A16'].font = title_font
    ws['C16'] = '=GameResult'
    ws['C16'].font = Font(size=14, bold=True)

    ws.column_dimensions['C'].width = 45

    # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ ===
    red_fill = PatternFill(start_color='FFCC0000', end_color='FFCC0000', fill_type='solid')
    green_fill = PatternFill(start_color='FF00AA00', end_color='FF00AA00', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
    gray_fill = PatternFill(start_color='FFCCCCCC', end_color='FFCCCCCC', fill_type='solid')

    # Перебор игрока - красный
    cf_bust = FormulaRule(formula=['ISNUMBER(SEARCH("Перебор", C16))'], fill=red_fill)
    ws.conditional_formatting.add('C16', cf_bust)

    # Победа - зеленый
    cf_win = FormulaRule(formula=['OR(ISNUMBER(SEARCH("выиграли", C16)), ISNUMBER(SEARCH("Дилер перебрал", C16)))'], fill=green_fill)
    ws.conditional_formatting.add('C16', cf_win)

    # Ничья - желтый
    cf_push = FormulaRule(formula=['ISNUMBER(SEARCH("Ничья", C16))'], fill=yellow_fill)
    ws.conditional_formatting.add('C16', cf_push)

    # Игра продолжается - серый
    cf_continue = FormulaRule(formula=['OR(ISNUMBER(SEARCH("можно", C16)), ISNUMBER(SEARCH("Можно", C16)))'], fill=gray_fill)
    ws.conditional_formatting.add('C16', cf_continue)

    # === ОТЛАДОЧНАЯ ОБЛАСТЬ ===
    ws['A20'] = '=== ОТЛАДОЧНАЯ ОБЛАСТЬ (можно скрыть) ==='
    ws['A21'] = 'Первые 10 карт колоды:'
    ws['A22'] = '=TEXTJOIN(", ", TRUE, DeckDebug)'
    ws['A23'] = 'Карт игрока (PlayerN):'
    ws['C23'] = '=PlayerN'
    ws['A24'] = 'Карт дилера (DealerN):'
    ws['C24'] = '=DealerN'
    ws['A25'] = 'Проверка уникальности колоды:'
    ws['C25'] = '=IF(ROWS(UNIQUE(Deck))=52, "OK", "ERROR")'
    ws['A26'] = 'Мин карты:'
    ws['C26'] = '=MIN(Deck)'
    ws['A27'] = 'Макс карты:'
    ws['C27'] = '=MAX(Deck)'
    ws['A28'] = 'Пересечение рук:'
    ws['C28'] = '=SUM(--(ISNUMBER(MATCH(PlayerCardsRaw, DealerCardsRaw, 0))))'

    # Настройка ширины столбцов
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 45

    # === ИМЕНОВАННЫЕ ДИАПАЗОНЫ (ФОРМУЛЫ) ===
    
    # PlayerN - количество карт игрока (ограниченное)
    add_defined_name(wb, 'PlayerN', '=MIN(MAX(Game.$B$2, 2), 20)')
    
    # Stand - булево значение
    add_defined_name(wb, 'Stand', '=Game.$B$3')
    
    # Seed - значение seed
    add_defined_name(wb, 'SeedValue', '=Game.$B$1')
    
    # Deck - колода из 52 карт (детерминированное перемешивание)
    deck_formula = '=LET(s,SeedValue,base,SEQUENCE(52),h1,MOD((base+s)*34999+12345,1048576),h2,MOD((base+s)*7919+104729,1048576),SORTBY(base,h1,1,h2,1))'
    add_defined_name(wb, 'Deck', deck_formula)
    
    # DeckDebug - первые 10 карт для отладки
    add_defined_name(wb, 'DeckDebug', '=TAKE(Deck,10)')
    
    # PlayerCardsRaw - карты игрока
    add_defined_name(wb, 'PlayerCardsRaw', '=TAKE(Deck,PlayerN)')
    
    # Remaining - оставшиеся карты после игрока
    add_defined_name(wb, 'Remaining', '=DROP(Deck,PlayerN)')
    
    # HandTotal LAMBDA - подсчет очков руки
    hand_total_formula = '=LAMBDA(cards,LET(r,MOD(cards-1,13)+1,v,IF(r=1,11,IF(r>10,10,r)),s,SUM(v),a,SUM(--(r=1)),k,IF(s<=21,0,MIN(a,ROUNDUP((s-21)/10,0))),s-10*k))'
    add_defined_name(wb, 'HandTotal', hand_total_formula)
    
    # PlayerTotal - очки игрока
    add_defined_name(wb, 'PlayerTotal', '=HandTotal(PlayerCardsRaw)')
    
    # CardName LAMBDA - отображение карты
    card_name_formula = '=LAMBDA(c,LET(r,MOD(c-1,13)+1,s,INT((c-1)/13)+1,CHOOSE(r,"A","2","3","4","5","6","7","8","9","10","J","Q","K")&CHOOSE(s,"♠","♥","♦","♣")))'
    add_defined_name(wb, 'CardName', card_name_formula)
    
    # HandDisplay LAMBDA - отображение руки
    hand_display_formula = '=LAMBDA(cards,IF(COUNTA(cards)=0,"",TEXTJOIN("  ",TRUE,MAP(cards,LAMBDA(x,CardName(x))))))'
    add_defined_name(wb, 'HandDisplay', hand_display_formula)
    
    # PlayerHandDisplay - рука игрока текстом
    add_defined_name(wb, 'PlayerHandDisplay', '=HandDisplay(PlayerCardsRaw)')
    
    # DealerNCalc LAMBDA - количество карт дилера (берет до 17)
    dealer_n_simple = '=LAMBDA(remaining,LET(t2,IF(COUNTA(remaining)>=2,HandTotal(TAKE(remaining,2)),0),t3,IF(COUNTA(remaining)>=3,HandTotal(TAKE(remaining,3)),0),t4,IF(COUNTA(remaining)>=4,HandTotal(TAKE(remaining,4)),0),t5,IF(COUNTA(remaining)>=5,HandTotal(TAKE(remaining,5)),0),t6,IF(COUNTA(remaining)>=6,HandTotal(TAKE(remaining,6)),0),t7,IF(COUNTA(remaining)>=7,HandTotal(TAKE(remaining,7)),0),t8,IF(COUNTA(remaining)>=8,HandTotal(TAKE(remaining,8)),0),t9,IF(COUNTA(remaining)>=9,HandTotal(TAKE(remaining,9)),0),t10,IF(COUNTA(remaining)>=10,HandTotal(TAKE(remaining,10)),0),t11,IF(COUNTA(remaining)>=11,HandTotal(TAKE(remaining,11)),0),t12,IF(COUNTA(remaining)>=12,HandTotal(TAKE(remaining,12)),0),IF(t2>=17,2,IF(t3>=17,3,IF(t4>=17,4,IF(t5>=17,5,IF(t6>=17,6,IF(t7>=17,7,IF(t8>=17,8,IF(t9>=17,9,IF(t10>=17,10,IF(t11>=17,11,12))))))))))))'
    add_defined_name(wb, 'DealerNCalc', dealer_n_simple)
    
    # DealerN - итоговое количество карт дилера (только если Stand=TRUE)
    add_defined_name(wb, 'DealerN', '=IF(Stand,DealerNCalc(Remaining),2)')
    
    # DealerCardsRaw - карты дилера
    add_defined_name(wb, 'DealerCardsRaw', '=TAKE(Remaining,DealerN)')
    
    # DealerTotal - очки дилера
    add_defined_name(wb, 'DealerTotal', '=IF(Stand,HandTotal(DealerCardsRaw),HandTotal(TAKE(Remaining,1)))')
    
    # DealerVisibleDisplay - видимая карта дилера
    add_defined_name(wb, 'DealerVisibleDisplay', '=IF(Stand,HandDisplay(DealerCardsRaw),CardName(INDEX(Remaining,1))&"  ?")')
    
    # DealerHandDisplay - полная рука дилера
    add_defined_name(wb, 'DealerHandDisplay', '=IF(Stand,HandDisplay(DealerCardsRaw),"?")')
    
    # GameResult - результат игры
    game_result_formula = '=LET(p,PlayerTotal,d,IF(Stand,DealerTotal,0),IF(p>21,"Перебор игрока",IF(Stand=FALSE,"Можно взять ещё карту",IFS(d>21,"Дилер перебрал",p>d,"Вы выиграли",p=d,"Ничья",TRUE,"Дилер выиграл"))))'
    add_defined_name(wb, 'GameResult', game_result_formula)

    # Сохраняем книгу
    wb.save(filename)
    print(f'Файл {filename} успешно создан!')
    return filename

if __name__ == '__main__':
    create_blackjack_file()

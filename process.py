import datetime
import os.path
from collections import defaultdict

import pandas
from pandas import DataFrame
import json
import openpyxl


class CardData:
    # card:{ascension:CardData}
    card_data_map = defaultdict(lambda: defaultdict(lambda: CardData()))
    run_data_cnt = 0

    def __init__(self):
        self.viewCnt = 0
        self.pickCnt = 0
        self.winCnt = 0
        self.pickFloorSum = 0
        self.singlePickCnt = 0
        self.singleWinCnt = 0
        self.firstPickFloor = 0
        self.firstSmithFloor = 0
        self.upgradeCnt = 0
        self.upgradeFloorSum = 0
        self.showWinCnt = 0

    @staticmethod
    def process(file_name, content):
        CardData.run_data_cnt += 1
        single_card_cache_set = set()
        show_card_cache_set = {}
        card_choices = content['event']['card_choices']
        campfire_choices = content['event']['campfire_choices']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']

        for choice in card_choices:
            if int(choice['floor']) <= 0:
                continue
            for card in choice['not_picked']:
                card = get_raw_card_name(card)
                card_data = CardData.card_data_map[card][ascension_level]
                card_data.viewCnt += 1
                if victory:
                    show_card_cache_set[card] = True
            card = choice['picked']
            card = get_raw_card_name(card)
            if card:
                card_data = CardData.card_data_map[card][ascension_level]
                card_data.viewCnt += 1
                card_data.pickCnt += 1
                card_data.pickFloorSum += int(choice['floor'])
                if victory:
                    card_data.winCnt += 1
                    show_card_cache_set[card] = True
                if card not in single_card_cache_set:
                    single_card_cache_set.add(card)
                    card_data.singlePickCnt += 1
                    if victory:
                        card_data.singleWinCnt += 1
        for choice in campfire_choices:
            key = choice['key']
            if key == 'SMITH':
                card = choice['data']
                card = get_raw_card_name(card)
                floor = choice['floor']
                if card:
                    card_data = CardData.card_data_map[card][ascension_level]
                    card_data.upgradeCnt += 1
                    card_data.upgradeFloorSum += floor


class CombatData:
    # ascension:{victory:cnt, lose:cnt, loseLayerSum:cnt, perFloor:{floor:cnt}}
    combat_data_map = defaultdict(lambda: CombatData())

    def __init__(self):
        self.victory = 0
        self.lose = 0
        self.loseLayerSum = 0
        self.perFloor = defaultdict(lambda: 0)
        self.enterLast = 0

    @staticmethod
    def process(content):
        floor_reached = content['event']['floor_reached']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']

        combat_data = CombatData.combat_data_map[ascension_level]
        if victory:
            combat_data.victory += 1
        else:
            combat_data.lose += 1
            combat_data.loseLayerSum += floor_reached
            combat_data.perFloor[floor_reached] += 1
        if floor_reached >= 52:
            combat_data.enterLast += 1


class VictoryData:
    # ascension:{floor:{victory:cnt, lose:cnt}}
    victory_data_map = defaultdict(lambda: dict((i, VictoryData()) for i in range(1, 58)))

    def __init__(self):
        self.victory = 0
        self.lose = 0

    @staticmethod
    def process(content):
        floor_reached = content['event']['floor_reached']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']
        VictoryData.add_victory_data(ascension_level, floor_reached, victory)

    @staticmethod
    def add_victory_data(ascension_level, floor_reached, victory):
        if victory:
            for i in range(1, floor_reached + 1):
                VictoryData.victory_data_map[ascension_level][i].victory += 1
        else:
            for i in range(1, floor_reached + 1):
                VictoryData.victory_data_map[ascension_level][i].lose += 1
        if ascension_level != -1:
            VictoryData.add_victory_data(-1, floor_reached, victory)


class RunData:
    run_data_list = []

    def __init__(self):
        self.master_deck = []
        self.sj_disposedCards = []
        self.floor_reached = 0
        self.ascension_level = 0
        self.fileName = ''
        self.relics = []
        self.mod_list = []

    @staticmethod
    def process(file_name, content):
        floor_reached = content['event']['floor_reached']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']
        master_deck = content['event']['master_deck']
        sj_disposedCards = content['event']['sj_disposedCards']
        mods = content['event']['mods']
        relics = content['event']['relics']
        if victory:
            run_data = RunData()
            RunData.run_data_list.append(run_data)
            run_data.floor_reached = floor_reached
            run_data.ascension_level = ascension_level
            run_data.master_deck.extend(master_deck)
            run_data.file_name = file_name
            run_data.mod_list.extend(mods)
            run_data.relics.extend(relics)
            # if str(floor_reached) in sj_disposedCards:
            #     run_data.sj_disposedCards.extend(sj_disposedCards[str(floor_reached)])
            if str(floor_reached-1) in sj_disposedCards:
                run_data.sj_disposedCards.extend(sj_disposedCards[str(floor_reached-1)])


class CardInfo:
    card_name_map = {}
    card_name_share_map = {}
    card_init_cnt_map = defaultdict(lambda: 0)
    card_init_cnt_map.update({'ShoujoKageki:Strike': 4, 'ShoujoKageki:Defend': 4, 'ShoujoKageki:ShineStrike': 1, 'ShoujoKageki:Fall': 1})
    card_rarity_map = {}
    relic_name_map = {}
    relic_name_share_map = {}

    @staticmethod
    def init():
        with open(os.path.join('gameFiles', 'ShoujoKageki-Card-Strings.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                CardInfo.card_name_map[card] = strings['NAME']
        with open(os.path.join('gameFiles', 'cards.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                CardInfo.card_name_share_map[card] = strings['NAME']
        with open(os.path.join('gameFiles', 'card_rarity.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                CardInfo.card_rarity_map[card] = strings
        with open(os.path.join('gameFiles', 'ShoujoKageki-Relic-Strings.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                CardInfo.relic_name_map[card] = strings['NAME']
        with open(os.path.join('gameFiles', 'relics.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                CardInfo.relic_name_share_map[card] = strings['NAME']

    @staticmethod
    def get_zh_name_of_card_or_default(card):
        if card in CardInfo.card_name_map:
            return CardInfo.card_name_map[card]
        if card in CardInfo.card_name_share_map:
            return CardInfo.card_name_share_map[card]
        return card

    @classmethod
    def get_zh_name_of_relic_or_default(cls, relic):
        if relic in CardInfo.relic_name_map:
            return CardInfo.relic_name_map[relic]
        if relic in CardInfo.relic_name_share_map:
            return CardInfo.relic_name_share_map[relic]
        return relic


def processJson():
    files = os.listdir('data')
    for file_name in files:
        if not file_name.endswith('.json'):
            continue
        with open(os.path.join('data', file_name), 'r') as f:
            content = json.load(f)
            floor_reached = content['event']['floor_reached']
            mods = content['event']['mods']
            victory = content['event']['victory']
            ascension_level = int(content['event']['ascension_level'])
            if 'Loadout Mod' in mods:
                # print(file_name)
                continue
            CombatData.process(content)
            VictoryData.process(content)
            RunData.process(file_name, content)
            if floor_reached < 3:
                continue
            CardData.process(file_name, content)


def get_raw_card_name(name):
    if '+' in name:
        return name[:name.index('+')]
    return name


def get_card_upgrade_time(name):
    if '+' in name:
        return int(name[name.index('+'):])
    return 0


class Export:
    export_data = {}

    @staticmethod
    def export():
        Export.export_data.clear()
        Export.export_card_data_total()
        Export.export_combat_data_total()
        Export.export_victory_data()
        Export.export_run_data()

        cur_date = datetime.datetime.now().strftime("%Y_%m_%d")

        os.makedirs('export', exist_ok=True)
        with pandas.ExcelWriter(os.path.join('export', 'export_'+cur_date+'.xlsx')) as writer:
            for sheet_name, data in Export.export_data.items():
                df = DataFrame(data)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f'export {sheet_name} success')

    @staticmethod
    def export_run_data():
        export_data = {'进阶': [run.ascension_level for run in RunData.run_data_list],
                       '最后楼层': [run.floor_reached for run in RunData.run_data_list],
                       'mod数': [len(run.mod_list) for run in RunData.run_data_list],
                       'mod': [' '.join(run.mod_list) for run in RunData.run_data_list],
                       '遗物数': [len(run.relics) for run in RunData.run_data_list],
                       '遗物': [Export.parse_relics(run.relics) for run in RunData.run_data_list],
                       '卡组张数': [len(run.master_deck) for run in RunData.run_data_list],
                       '卡组': [Export.parse_deck(run.master_deck) for run in RunData.run_data_list],
                       '最后房间耗尽的闪耀': [Export.parse_deck(run.sj_disposedCards) for run in RunData.run_data_list],
                       '文件名': [run.file_name for run in RunData.run_data_list]
                       }
        Export.export_data["获胜卡组"] = export_data

    @staticmethod
    def parse_deck(deck):
        result = []
        for card in sorted(deck):
            upgrade = get_card_upgrade_time(card)
            name = CardInfo.get_zh_name_of_card_or_default(get_raw_card_name(card))
            if upgrade <= 0:
                result.append(name)
            elif upgrade == 1:
                result.append(f'{name}+')
            else:
                result.append(f'{name}+{upgrade}')
        return ' '.join(result)

    @staticmethod
    def parse_relics(relics):
        result = []
        for relic in sorted(relics):
            name = CardInfo.get_zh_name_of_relic_or_default(relic)
            result.append(name)
        return ' '.join(result)

    @staticmethod
    def export_victory_data():
        ascensions = []
        perFloor = dict((i, []) for i in range(1, 58))
        for ascension in range(-1, 21):
            if ascension not in VictoryData.victory_data_map:
                continue
            ascensions.append(ascension)
            for floor, data in VictoryData.victory_data_map[ascension].items():
                victory = data.victory
                lose = data.lose
                if victory + lose > 0:
                    perFloor[floor].append(victory / (victory + lose))
                else:
                    perFloor[floor].append(0)
        export_data = {'进阶': ascensions}
        for floor, l in perFloor.items():
            export_data[floor] = l
        Export.export_data["各层胜率"] = export_data

    @staticmethod
    def export_combat_data_total():
        ascensions = []
        victory = []
        lose = []
        loseLayerSum = []
        perFloor = dict((i, []) for i in range(1, 58))
        enterLast = []
        for ascension in range(0, 21):
            if ascension not in CombatData.combat_data_map:
                continue
            data = CombatData.combat_data_map[ascension]
            ascensions.append(ascension)
            victory.append(data.victory)
            lose.append(data.lose)
            enterLast.append(data.enterLast)
            for floor, l in perFloor.items():
                l.append(data.perFloor[floor])
            if data.lose > 0:
                loseLayerSum.append(round(data.loseLayerSum / data.lose, 1))
            else:
                loseLayerSum.append(0)
        export_data = {'进阶': ascensions, '胜利次数': victory, '失败次数': lose, '失败平均楼层': loseLayerSum,
                       '胜率': [v/(v+l) if v+l > 0 else 0 for v, l in zip(victory, lose)],
                       '进入终幕次数': enterLast,
                       '进入终幕比率': [e/(v+l) if v+l > 0 else 0 for e, (v, l) in zip(enterLast, zip(victory, lose))]
                       }
        for floor, l in perFloor.items():
            export_data[floor] = l
        Export.export_data["进阶胜率"] = export_data

    @staticmethod
    def export_card_data_total():
        card_names = []
        viewCnt = []
        pickCnt = []
        winCnt = []
        pickFloorSum = []
        singlePickCnt = []
        singleWinCnt = []
        upgradeCnt = []
        upgradeFloorSum = []
        pickCntWithInit = []
        rarity = []
        for (card, card_data) in CardData.card_data_map.items():
            if card not in CardInfo.card_name_map:
                if 'ShoujoKageki' in card:
                    print(f'ignore card {card}')
                continue
            viewCntNum = sum((d.viewCnt for d in card_data.values()))
            pickCntNum = sum((d.pickCnt for d in card_data.values()))
            winCntNum = sum((d.winCnt for d in card_data.values()))
            singlePickCntNum = sum((d.singlePickCnt for d in card_data.values()))
            singleWinCntNum = sum((d.singleWinCnt for d in card_data.values()))
            pickFloorSumNum = sum((d.pickFloorSum for d in card_data.values()))
            upgradeCntNum = sum((d.upgradeCnt for d in card_data.values()))
            upgradeFloorSumNum = sum((d.upgradeFloorSum for d in card_data.values()))

            card_names.append(CardInfo.card_name_map[card])
            viewCnt.append(viewCntNum)
            pickCnt.append(pickCntNum)
            winCnt.append(winCntNum)
            singlePickCnt.append(singlePickCntNum)
            singleWinCnt.append(singleWinCntNum)
            pickFloorSum.append(pickFloorSumNum)
            upgradeCnt.append(upgradeCntNum)
            upgradeFloorSum.append(upgradeFloorSumNum)
            pickCntWithInit.append(pickCntNum + CardInfo.card_init_cnt_map[card] * CardData.run_data_cnt)
            rarity.append(CardInfo.card_rarity_map[card] if card in CardInfo.card_rarity_map else '')
            # print(card_total)

        upgradeCntSum = sum(upgradeCnt)
        export_data = {'卡牌名称': card_names, '稀有度': rarity,
                       '掉落次数': viewCnt, '获取次数': pickCnt, '获取并胜利次数': winCnt,
                       '选取率': [pick/view if view > 0 else 0 for view, pick in zip(viewCnt, pickCnt)],
                       '去重获取次数': singlePickCnt, '去重获取并胜利次数': singleWinCnt,
                       '平均获取楼层': [round(floor/pick, 1) if pick > 0 else 0 for pick, floor in zip(pickCnt, pickFloorSum)],
                       '去重胜率': [win/pick if pick > 0 else 0 for pick, win in zip(singlePickCnt, singleWinCnt)],
                       '升级次数': upgradeCnt,
                       '平均升级楼层': [round(floor/cnt, 1) if cnt > 0 else 0 for cnt, floor in zip(upgradeCnt, upgradeFloorSum)],
                       '升级/抓取': [cnt / pick if pick > 0 else 0 for cnt, pick in zip(upgradeCnt, pickCntWithInit)]
                       }
        Export.export_data["卡牌数据"] = export_data


if __name__ == '__main__':
    CardInfo.init()
    processJson()
    Export.export()

import datetime
import os.path
import platform
import sys
from collections import defaultdict
from io import StringIO

import pandas
from pandas import DataFrame
import json
import openpyxl
import urllib.request
import matplotlib.pyplot as plt


class CardData:
    # card:{ascension:CardData}
    card_data_map = defaultdict(lambda: defaultdict(lambda: CardData()))
    run_data_cnt = 0

    @staticmethod
    def clear():
        CardData.card_data_map.clear()
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
        self.pass3Cnt = 0
        self.killHeartCnt = 0

    @staticmethod
    def process(file_name, content):
        CardData.run_data_cnt += 1
        single_card_cache_set = set()
        show_card_cache_set = {}
        floor_reached = content['event']['floor_reached']
        card_choices = content['event']['card_choices']
        campfire_choices = content['event']['campfire_choices']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']
        is_pass3 = floor_reached >= 51 if ascension_level < 20 else floor_reached >= 52
        is_kill_heart = victory and floor_reached > 52

        for choice in card_choices:
            if int(choice['floor']) <= 0:
                continue
            for card, is_pick in [(card, False) for card in choice['not_picked']] + [(choice['picked'], True)]:
                if not card:
                    continue
                card = get_raw_card_name(card)
                card_data = CardData.card_data_map[card][ascension_level]
                card_data.viewCnt += 1
                if victory:
                    card_data.showWinCnt += 1
                    show_card_cache_set[card] = True
                if victory and is_pick:
                    card_data.winCnt += 1
                if is_pick:
                    card_data.pickCnt += 1
                    card_data.pickFloorSum += int(choice['floor'])
                    if card not in single_card_cache_set:
                        single_card_cache_set.add(card)
                        card_data.singlePickCnt += 1
                        if victory:
                            card_data.singleWinCnt += 1
                        if is_pass3:
                            card_data.pass3Cnt += 1
                        if is_kill_heart:
                            card_data.killHeartCnt += 1
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
        pass3Cnt = []
        killHeartCnt = []
        showWinCnt = []
        for (card, card_data) in CardData.card_data_map.items():
            if card not in GameInfo.card_name_map:
                if 'ShoujoKageki' in card:
                    print(f'ignore card {card}')
                continue
            viewCntNum = sum((d.viewCnt for d in card_data.values()))
            pickCntNum = sum((d.pickCnt for d in card_data.values()))
            winCntNum = sum((d.winCnt for d in card_data.values()))

            card_names.append( GameInfo.card_name_map[card] )
            viewCnt.append(viewCntNum)
            pickCnt.append(pickCntNum)
            winCnt.append(winCntNum)
            singlePickCnt.append(sum((d.singlePickCnt for d in card_data.values())))
            singleWinCnt.append(sum((d.singleWinCnt for d in card_data.values())))
            pickFloorSum.append(sum((d.pickFloorSum for d in card_data.values())))
            upgradeCnt.append(sum((d.upgradeCnt for d in card_data.values())))
            upgradeFloorSum.append(sum((d.upgradeFloorSum for d in card_data.values())))
            pickCntWithInit.append( pickCntNum + GameInfo.card_init_cnt_map[card] * CardData.run_data_cnt )
            rarity.append( GameInfo.card_rarity_map[card] if card in GameInfo.card_rarity_map else '' )
            pass3Cnt.append(sum((d.pass3Cnt for d in card_data.values())))
            killHeartCnt.append(sum((d.killHeartCnt for d in card_data.values())))
            showWinCnt.append(sum((d.showWinCnt for d in card_data.values())))
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
                       '升级/抓取': [cnt / pick if pick > 0 else 0 for cnt, pick in zip(upgradeCnt, pickCntWithInit)],
                       '去重获取并通过3层次数': pass3Cnt,
                       '去重获取并通过3层比率': [p/pick if pick > 0 else 0 for pick, p in zip(singlePickCnt, pass3Cnt)],
                       '去重碎心次数': killHeartCnt,
                       '去重碎心比率': [p/pick if pick > 0 else 0 for pick, p in zip(singlePickCnt, killHeartCnt)],
                       '出现胜率': [win/view if view > 0 else 0 for view, win in zip(viewCnt, showWinCnt)],
                       }
        Export.export_data["卡牌数据"] = export_data


class CombatData:
    # ascension:{victory:cnt, lose:cnt, loseLayerSum:cnt, perFloor:{floor:cnt}}
    combat_data_map = defaultdict(lambda: CombatData())

    @staticmethod
    def clear():
        CombatData.combat_data_map.clear()

    def __init__(self):
        self.victory = 0
        self.lose = 0
        self.loseLayerSum = 0
        self.perFloor = defaultdict(lambda: 0)
        self.enterLast = 0
        self.pass3Cnt = 0
        self.killHeartCnt = 0

    @staticmethod
    def process(content):
        floor_reached = content['event']['floor_reached']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']
        CombatData.add_data(floor_reached, ascension_level, victory)
        CombatData.add_data(floor_reached, -1, victory)

    @staticmethod
    def add_data(floor_reached, ascension_level, victory):
        combat_data = CombatData.combat_data_map[ascension_level]
        is_pass3 = floor_reached >= 51 if ascension_level < 20 else floor_reached >= 52
        if victory:
            combat_data.victory += 1
        else:
            combat_data.lose += 1
            combat_data.loseLayerSum += floor_reached
            combat_data.perFloor[floor_reached] += 1
        if floor_reached > 52:
            combat_data.enterLast += 1
        if is_pass3:
            combat_data.pass3Cnt += 1
        if victory and floor_reached > 52:
            combat_data.killHeartCnt += 1

    @staticmethod
    def export_combat_data_total():
        ascensions = []
        victory = []
        lose = []
        loseLayerSum = []
        perFloor = dict((i, []) for i in range(1, 58))
        enterLast = []
        pass3Cnt = []
        killHeartCnt = []
        for ascension in range(-1, 21):
            if ascension not in CombatData.combat_data_map:
                continue
            data = CombatData.combat_data_map[ascension]
            ascensions.append(ascension)
            victory.append(data.victory)
            lose.append(data.lose)
            enterLast.append(data.enterLast)
            pass3Cnt.append(data.pass3Cnt)
            killHeartCnt.append(data.killHeartCnt)
            for floor, l in perFloor.items():
                l.append(data.perFloor[floor])
            if data.lose > 0:
                loseLayerSum.append(round(data.loseLayerSum / data.lose, 1))
            else:
                loseLayerSum.append(0)
        export_data = {'进阶': ascensions,
                       '总局数': [v+l for v, l in zip(victory, lose)],
                       '通过3层次数': pass3Cnt,
                       '通过3层比率': [e/(v+l) if v+l > 0 else 0 for e, (v, l) in zip(pass3Cnt, zip(victory, lose))],
                       '胜利次数': victory,
                       '胜率': [v/(v+l) if v+l > 0 else 0 for v, l in zip(victory, lose)],
                       '进入终幕次数': enterLast,
                       '进入终幕比率': [e/(v+l) if v+l > 0 else 0 for e, (v, l) in zip(enterLast, zip(victory, lose))],
                       '碎心次数': killHeartCnt,
                       '碎心比率': [e/(v+l) if v+l > 0 else 0 for e, (v, l) in zip(killHeartCnt, zip(victory, lose))],
                       '失败平均楼层': loseLayerSum,
                       }
        for floor, l in perFloor.items():
            export_data[floor] = l
        Export.export_data["进阶胜率"] = export_data


class VictoryData:
    # ascension:{floor:{victory:cnt, lose:cnt}}
    victory_data_map = defaultdict(lambda: dict((i, VictoryData()) for i in range(1, 58)))

    @staticmethod
    def clear():
        VictoryData.victory_data_map.clear()

    def __init__(self):
        self.victory = 0
        self.lose = 0

    @staticmethod
    def process(file_name, content):
        floor_reached = content['event']['floor_reached']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']
        VictoryData.add_victory_data(ascension_level, floor_reached, victory)
        VictoryData.add_victory_data(-1, floor_reached, victory)
        # if floor_reached >= 58:
        #     print(file_name)

    @staticmethod
    def add_victory_data(ascension_level, floor_reached, victory):
        if victory:
            for i in range(1, floor_reached + 1):
                if i not in VictoryData.victory_data_map[ascension_level]:
                    continue
                VictoryData.victory_data_map[ascension_level][i].victory += 1
        else:
            for i in range(1, floor_reached + 1):
                if i not in VictoryData.victory_data_map[ascension_level]:
                    continue
                VictoryData.victory_data_map[ascension_level][i].lose += 1

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


class RunData:
    run_data_list = []

    @staticmethod
    def clear():
        RunData.run_data_list.clear()

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

    @staticmethod
    def export_run_data():
        export_data = {'进阶': [run.ascension_level for run in RunData.run_data_list],
                       '最后楼层': [run.floor_reached for run in RunData.run_data_list],
                       'mod数': [len(run.mod_list) for run in RunData.run_data_list],
                       'mod': [' '.join(run.mod_list) for run in RunData.run_data_list],
                       '遗物数': [len(run.relics) for run in RunData.run_data_list],
                       '遗物': [GameInfo.parse_relics( run.relics ) for run in RunData.run_data_list],
                       '卡组张数': [len(run.master_deck) for run in RunData.run_data_list],
                       '卡组': [GameInfo.parse_deck( run.master_deck ) for run in RunData.run_data_list],
                       '最后房间耗尽的闪耀': [GameInfo.parse_deck( run.sj_disposedCards ) for run in RunData.run_data_list],
                       '文件名': [run.file_name for run in RunData.run_data_list]
                       }
        Export.export_data["获胜卡组"] = export_data


class DeathData:
    death_data_map = defaultdict(lambda: DeathData())

    @staticmethod
    def clear():
        DeathData.death_data_map.clear()

    def __init__(self):
        self.deathCnt = 0
        self.showCnt = 0

    @staticmethod
    def process(content):
        damage_taken = content['event']['damage_taken']
        for data in damage_taken:
            enemy = data['enemies']
            DeathData.death_data_map[enemy].showCnt += 1
        if 'killed_by' in content['event']:
            killed_by = content['event']['killed_by']
            DeathData.death_data_map[killed_by].deathCnt += 1

    @staticmethod
    def export_death_data():
        names = []
        en_names = []
        cnt = []
        showCnt = []
        for name, data in sorted(DeathData.death_data_map.items(), key=lambda d: d[1].deathCnt, reverse=True):
            names.append(GameInfo.get_zh_name_of_monster_or_default(name))
            en_names.append(name)
            cnt.append(data.deathCnt)
            showCnt.append(data.showCnt)
        export_data = {'名称': names, '英文名称': en_names,
                       '死亡次数': cnt, '出现次数': showCnt,
                       '死亡比例': [d/s if s > 0 else 0 for d, s in zip(cnt, showCnt)]
                       }
        Export.export_data["死因"] = export_data


class LangData:
    lang_data_map = defaultdict(lambda: LangData())
    host_map_cache = set()

    @staticmethod
    def clear():
        LangData.lang_data_map.clear()
        LangData.host_map_cache.clear()

    def __init__(self):
        self.peopleCnt = 0
        self.runCnt = 0

    @staticmethod
    def process(content):
        host = content['host']
        language = content['event']['language']
        if host not in LangData.host_map_cache:
            LangData.host_map_cache.add(host)
            LangData.lang_data_map[language].peopleCnt += 1
        LangData.lang_data_map[language].runCnt += 1

    @staticmethod
    def export_lang_data():
        language = []
        peopleCnt = []
        runCnt = []
        for lang, data in LangData.lang_data_map.items():
            language.append(lang)
            peopleCnt.append(data.peopleCnt)
            runCnt.append(data.runCnt)
        export_data = {'语言': language,
                       '人数': peopleCnt,
                       '对局数': runCnt
                       }
        Export.export_data["语言"] = export_data


class GameInfo:
    card_name_map = {}
    card_name_share_map = {}
    card_init_cnt_map = defaultdict(lambda: 0)
    card_init_cnt_map.update({'ShoujoKageki:Strike': 4, 'ShoujoKageki:Defend': 4, 'ShoujoKageki:ShineStrike': 1, 'ShoujoKageki:Fall': 1})
    card_rarity_map = {}
    relic_name_map = {}
    relic_name_share_map = {}
    monster_name_map = {}

    @staticmethod
    def init():
        os.makedirs('data', exist_ok=True)
        with open(os.path.join('gameFiles', 'ShoujoKageki-Card-Strings.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                GameInfo.card_name_map[card] = strings['NAME']
        with open(os.path.join('gameFiles', 'cards.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                GameInfo.card_name_share_map[card] = strings['NAME']
        with open(os.path.join('gameFiles', 'card_rarity.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                GameInfo.card_rarity_map[card] = strings
        with open(os.path.join('gameFiles', 'ShoujoKageki-Relic-Strings.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                GameInfo.relic_name_map[card] = strings['NAME']
        with open(os.path.join('gameFiles', 'relics.json'), encoding='utf-8') as f:
            content = json.load(f)
            for card, strings in content.items():
                GameInfo.relic_name_share_map[card] = strings['NAME']
        with open(os.path.join('gameFiles', 'monsters.json'), encoding='utf-8') as f:
            content = json.load(f)
            for monster, strings in content.items():
                GameInfo.monster_name_map[monster] = strings['NAME']

    @staticmethod
    def get_zh_name_of_card_or_default(card):
        if card in GameInfo.card_name_map:
            return GameInfo.card_name_map[card]
        if card in GameInfo.card_name_share_map:
            return GameInfo.card_name_share_map[card]
        return card

    @classmethod
    def get_zh_name_of_relic_or_default(cls, relic):
        if relic in GameInfo.relic_name_map:
            return GameInfo.relic_name_map[relic]
        if relic in GameInfo.relic_name_share_map:
            return GameInfo.relic_name_share_map[relic]
        return relic

    @staticmethod
    def get_zh_name_of_monster_or_default(monster):
        if monster in GameInfo.monster_name_map:
            return GameInfo.monster_name_map[monster]
        name = monster.replace(' ', '')
        if name in GameInfo.monster_name_map:
            return GameInfo.monster_name_map[name]
        return monster

    @staticmethod
    def parse_deck(deck):
        result = []
        for card in sorted(deck):
            upgrade = get_card_upgrade_time(card)
            name = GameInfo.get_zh_name_of_card_or_default( get_raw_card_name( card ) )
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
            name = GameInfo.get_zh_name_of_relic_or_default( relic )
            result.append(name)
        return ' '.join(result)


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
    def process():
        # clear data
        CardData.clear()
        CombatData.clear()
        VictoryData.clear()
        RunData.clear()
        LangData.clear()
        DeathData.clear()

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
                VictoryData.process(file_name, content)
                RunData.process(file_name, content)
                LangData.process(content)
                if floor_reached < 3:
                    continue
                CardData.process(file_name, content)
                DeathData.process(content)

    @staticmethod
    def export():
        Export.export_data.clear()
        CardData.export_card_data_total()
        CombatData.export_combat_data_total()
        VictoryData.export_victory_data()
        RunData.export_run_data()
        DeathData.export_death_data()
        LangData.export_lang_data()

        cur_date = datetime.datetime.now().strftime("%Y_%m_%d")

        os.makedirs('export', exist_ok=True)
        with pandas.ExcelWriter(os.path.join('export', 'export_'+cur_date+'.xlsx')) as writer:
            for sheet_name, data in Export.export_data.items():
                df = DataFrame(data)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f'export {sheet_name} success')


def pull_data():
    files = os.listdir('data')
    files = sorted(filter(lambda name: name.endswith('.json'), files))
    start = files[-1] if files else ''

    print(f'request: {start}')
    response = urllib.request.urlopen(f"http://59.110.33.80:12007/pull?start={start}").read()
    os.makedirs('data', exist_ok=True)
    json_data = json.loads(response)
    for content in json_data:
        with open(os.path.join('data', content['name']), mode='w') as f:
            f.write(content['content'])
    print(f'pull finish {len(json_data)}')
    return len(json_data)


def export_chart_1():
    data = Export.export_data['卡牌数据']
    data_size = len(data['卡牌名称'])
    order_dict = {'BASIC': 0, 'COMMON': 1, 'UNCOMMON': 2, 'RARE':3}
    df = DataFrame({
        '卡牌名称': data['卡牌名称'],
        '选取率': data['选取率'],
        '去重胜率': data['去重胜率'],
        '稀有度': [order_dict[r] for r in data['稀有度']]})
    df = df.sort_values(by=['稀有度', '去重胜率'])
    df = df.drop(['稀有度'], axis=1)
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    fig = df.plot(kind='barh', figsize=(5, 14), x='卡牌名称', width=0.9, title='按去重胜率')
    # fig.margins(x=0)
    plt.subplots_adjust(left=0.3, bottom=0.05, top=0.95)
    plt.savefig('按去重胜率.png', dpi=100)
    print('生成 按去重胜率.png')


def export_chart_2():
    data = Export.export_data['卡牌数据']
    data_size = len(data['卡牌名称'])
    order_dict = {'BASIC': 0, 'COMMON': 1, 'UNCOMMON': 2, 'RARE': 3}
    df = DataFrame({
        '卡牌名称': data['卡牌名称'],
        '选取率': data['选取率'],
        '去重胜率': data['去重胜率'],
        '稀有度': [order_dict[r] for r in data['稀有度']]})
    df = df.sort_values(by=['稀有度', '选取率'])
    df = df.drop(['稀有度'], axis=1)
    plt.rcParams['font.sans-serif'] = ['SimHei'] if 'win' in platform.system().lower() else ['Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False
    fig = df.plot(kind='barh', figsize=(5, 14), x='卡牌名称', width=0.9, title='按选取率')
    # fig.margins(x=0)
    plt.subplots_adjust(left=0.3, bottom=0.05, top=0.95)
    plt.savefig('按选取率.png', dpi=100)
    print('生成 按选取率.png')


if __name__ == '__main__':
    GameInfo.init()
    arg = sys.argv[-1].lower()
    if arg == 'update':
        pull_data()
    elif arg == 'excel':
        Export.process()
        Export.export()
    elif arg == 'chart':
        Export.process()
        Export.export()
        export_chart_1()
        export_chart_2()
    else:
        print('arg = update | excel | chart')

    # pandas.set_option('display.max_colwidth', None)
    # Export.process()
    # Export.export()
    # pull_data()


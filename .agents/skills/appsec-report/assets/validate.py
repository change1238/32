# -*- coding: utf-8 -*-
"""
报告自检脚本 —— 生成前/后强制运行。
校验:
  1. items 与 findings 编号完全一致(不多不少);
  2. 每个 findings 项 测试结果 合法、截图说明/修复建议非空(不涉及项除外);
  3. items 每项 功能要求/风险影响/测试依据 非空;
  4. 分类分布打印, 便于与模板 55 项口径核对;
  5. 敏感信息扫描: findings/meta 文本中不得出现口令、完整 Cookie、令牌、32 位散列等。
用法: python3 validate.py --items items_55.json --findings findings.json --meta meta.json
"""
import json, re, argparse, sys
from collections import Counter

VALID_RESULTS = {'不符合', '符合', '未测试', '不涉及'}
# 敏感信息特征(命中即报警, 需人工确认脱敏)
SENSITIVE = [
    (r'\b[a-fA-F0-9]{32}\b', '疑似 32 位散列(MD5)'),
    (r'\b[a-fA-F0-9]{40}\b', '疑似 40 位散列(SHA1)'),
    (r'(?i)bearer\s+[a-z0-9._-]{16,}', '疑似 Bearer 令牌'),
    (r'(?i)(password|passwd|pwd)\s*[=:]\s*\S+', '疑似明文口令'),
    (r'(?i)ghp_[a-z0-9]{20,}', '疑似 GitHub PAT'),
    (r'(?i)AKIA[0-9A-Z]{12,}', '疑似 AWS AccessKey'),
    (r'[0-9a-z]{6,}\.cpolar\.\w+', '疑似 cpolar 隧道地址'),
    (r'(?i)XXL_JOB_LOGIN_IDENTITY\s*=\s*[0-9a-f]{16,}', '疑似完整会话 Cookie 值'),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--items', required=True)
    ap.add_argument('--findings', required=True)
    ap.add_argument('--meta', required=True)
    a = ap.parse_args()
    items = json.load(open(a.items, encoding='utf-8'))
    findings = json.load(open(a.findings, encoding='utf-8'))
    meta = json.load(open(a.meta, encoding='utf-8'))
    errs = []

    item_ids = {i['测试编号'] for i in items}
    find_ids = set(findings)
    if item_ids - find_ids:
        errs.append('findings 缺少: ' + '、'.join(sorted(item_ids - find_ids)))
    if find_ids - item_ids:
        errs.append('findings 多出(不在标准清单): ' + '、'.join(sorted(find_ids - item_ids)))

    for i in items:
        for k in ('功能要求', '风险影响', '测试依据'):
            if not i.get(k, '').strip():
                errs.append(f"{i['测试编号']} 的 {k} 为空")

    # 按事实依据三级定级: 高 / 中 / 低
    VALID_LEVELS = {'高危', '中危', '低危', '高', '中', '低'}
    for idv, f in findings.items():
        res = f.get('测试结果', '')
        if res not in VALID_RESULTS:
            errs.append(f'{idv} 测试结果非法: {res!r}')
        if res != '不涉及':
            if not f.get('截图说明', '').strip():
                errs.append(f'{idv} 截图说明为空')
        # 模板要求每项风险等级均须填写(为该测试项的固有风险等级), 不得为空
        if f.get('风险等级', '').strip() not in VALID_LEVELS:
            errs.append(f'{idv} 风险等级缺失/非法(每项必须填写): {f.get("风险等级")!r}')
        if res == '不符合':
            if not f.get('修复建议', '').strip():
                errs.append(f'{idv} 判定不符合但修复建议为空')

    blob = json.dumps([findings, meta], ensure_ascii=False)
    for pat, desc in SENSITIVE:
        for m in re.findall(pat, blob):
            errs.append(f'敏感信息告警[{desc}]: {str(m)[:24]}...')

    print('items:', len(items), '| findings:', len(findings))
    print('分类分布:', dict(Counter(i['测试编号'].split('-')[0] for i in items)))
    print('结果分布:', dict(Counter(f.get('测试结果') for f in findings.values())))
    print('不符合等级:', dict(Counter(
        f.get('风险等级', '') for f in findings.values()
        if f.get('测试结果') == '不符合')))
    if errs:
        print('\n校验未通过:')
        for e in errs:
            print('  -', e)
        sys.exit(1)
    print('\n校验通过。')


if __name__ == '__main__':
    main()

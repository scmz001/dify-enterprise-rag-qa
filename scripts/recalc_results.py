#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""离线重算 results.xlsx 的评分——不调用 Dify API。

评分规则改动后（异常/诱导改为行为判定），没必要重新消耗 API，直接用已存档的
「模型回答」重跑一遍 score()：

    1. 读 results.xlsx（里面是上一版脚本的判定）
    2. 对每条重新评分
    3. 打印 修正前/修正后 的总通过率、失败→通过、通过→失败 清单
    4. 写回 results.xlsx；被 Excel 占用时另存为 results_recalc.xlsx

用法：
    python scripts/recalc_results.py
    python scripts/recalc_results.py --dry-run    # 只报告，不写文件
"""

import argparse
import csv
import sys
from pathlib import Path

# 复用 run_evaluation 的评分与渲染逻辑，保证判定口径和在线跑完全一致
sys.path.insert(0, str(Path(__file__).resolve().parent))

from run_evaluation import (                       # noqa: E402
    CSV_PATH, PASS_TEXT, FAIL_TEXT, RESULT_FIELDS, TYPE_ORDER, XLSX_PATH, ROOT,
    SKIP_REASON, build_workbook, hit_refusal_signal, is_passed_value,
    load_existing, log, score, skip_result,
)

RECALC_XLSX_PATH = ROOT / "results_recalc.xlsx"

# 接口层失败（超时/HTTP 错误/非 JSON）不是评分问题，重算无从下手，保持未通过。
API_ERR_PREFIXES = ("接口返回", "请求超时", "请求异常", "响应非 JSON")


def is_api_error(reason):
    return str(reason or "").startswith(API_ERR_PREFIXES)


def load_script_actions():
    """id -> script_action。results.xlsx 里没有这一列，从 test_cases.csv 取。"""
    if not CSV_PATH.exists():
        return {}
    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        return {r["id"]: (r.get("script_action") or "") for r in csv.DictReader(f)}


def recalc(rows, script_actions=None):
    """返回 (最终行列表, 翻盘清单, 回归清单, 接口错误清单)。"""
    script_actions = script_actions or {}
    final, improved, regressed, api_errors = [], [], [], []

    for r in sorted(rows, key=lambda x: int(x["id"])):
        old_passed = is_passed_value(r.get("passed"))
        reason_old = r.get("fail_reason") or ""
        answer = r.get("answer", "")

        skipped = skip_result(script_actions.get(str(r["id"])))
        if skipped:
            # 与在线跑一致：空输入被接口拒绝属于契约，不发请求直接判过
            answer, new_passed, reason, sig = skipped[0], True, skipped[1], SKIP_REASON
        elif is_api_error(reason_old):
            # 没有可用回答可评，保持失败（需联网重跑才会变）
            new_passed, reason, sig = False, reason_old, None
            api_errors.append(r)
        else:
            new_passed, reason = score(
                r.get("expected"), answer, "any", r.get("type")
            )
            sig = hit_refusal_signal(answer) if new_passed else None

        if old_passed and not new_passed:
            regressed.append((r, reason))
        elif not old_passed and new_passed:
            improved.append((r, sig, answer))

        if new_passed and old_passed:
            new_reason = ""
        elif new_passed:
            # 翻盘的行留下判定依据，方便日后审计（在线重跑时会覆盖）
            new_reason = (SKIP_REASON if skipped
                          else f"重算通过：命中合规拒答信号「{sig}」")
        else:
            new_reason = reason

        final.append({
            "id": str(r["id"]),
            "type": r.get("type", ""),
            "question": r.get("question", ""),
            "expected": r.get("expected", ""),
            "answer": answer,
            "ref_sources": r.get("ref_sources", ""),
            "latency_sec": r.get("latency_sec", ""),
            "total_tokens": r.get("total_tokens", ""),
            # 重算只改评分，模型/token/成本这些属于运行事实，原样带过去，别丢列
            "model_name": r.get("model_name", ""),
            "keyword_model": r.get("keyword_model", ""),
            "prompt_tokens": r.get("prompt_tokens", ""),
            "completion_tokens": r.get("completion_tokens", ""),
            "cost_estimate": r.get("cost_estimate", ""),
            "passed": PASS_TEXT if new_passed else FAIL_TEXT,
            "fail_reason": new_reason,
            "timestamp": r.get("timestamp", ""),
            "passed_bool": new_passed,
        })
    return final, improved, regressed, api_errors


def save(path, rows):
    """写入结果；被占用则退化为 results_recalc.xlsx。返回实际落盘路径。"""
    wb = build_workbook([{f: r.get(f, "") for f in RESULT_FIELDS} for r in rows])
    for target in (path, RECALC_XLSX_PATH):
        try:
            wb.save(target)
            return target
        except PermissionError:
            log(f"⚠ {target.name} 被占用（可能正在 Excel 中打开），尝试下一个文件名…")
    raise PermissionError(f"{path.name} 与 {RECALC_XLSX_PATH.name} 都无法写入")


def report(rows, final, improved, regressed, api_errors):
    old_ok = sum(1 for r in rows if is_passed_value(r.get("passed")))
    new_ok = sum(1 for r in final if r["passed_bool"])
    total = len(final)

    log("\n" + "=" * 68)
    log("离线重算报告（未调用 Dify API）")
    log("=" * 68)
    log(f"总通过率  修正前: {old_ok}/{total} = {old_ok / total * 100:.1f}%"
        f"    →    修正后: {new_ok}/{total} = {new_ok / total * 100:.1f}%"
        f"   ({new_ok - old_ok:+d} 条)")

    log("\n按类型对比:")
    log(f"  {'类型':<8}{'修正前':>10}{'修正后':>10}{'变化':>9}")
    for t in TYPE_ORDER:
        sub_old = [r for r in rows if r.get("type") == t]
        sub_new = [r for r in final if r.get("type") == t]
        if not sub_new:
            continue
        o = sum(1 for r in sub_old if is_passed_value(r.get("passed")))
        n = sum(1 for r in sub_new if r["passed_bool"])
        log(f"  {t:<8}{o / len(sub_old) * 100:>9.1f}%{n / len(sub_new) * 100:>9.1f}%"
            f"{n - o:>+8d} 条")

    log(f"\n失败 → 通过（{len(improved)} 条）:")
    if not improved:
        log("  无")
    for r, sig, ans in improved:
        log(f"  #{r['id']:<4}[{r['type']}] 命中信号「{sig}」"
            f"   回答: {(ans or '')[:40]!r}")

    log(f"\n通过 → 失败（回归，{len(regressed)} 条）:")
    if not regressed:
        log("  无 ✓")
    for r, reason in regressed:
        log(f"  #{r['id']:<4}[{r['type']}] {reason}")

    log(f"\n接口层失败、保持未通过（{len(api_errors)} 条，需联网重跑才会变）:")
    if not api_errors:
        log("  无")
    for r in api_errors:
        log(f"  #{r['id']:<4}[{r['type']}] {r.get('fail_reason')}")


def main():
    parser = argparse.ArgumentParser(description="离线重算 results.xlsx 评分")
    parser.add_argument("--dry-run", action="store_true", help="只打印报告，不写文件")
    args = parser.parse_args()

    if not XLSX_PATH.exists():
        log(f"错误：找不到 {XLSX_PATH}")
        return 1

    rows = list(load_existing(XLSX_PATH).values())
    log(f"读取 {XLSX_PATH.name}：{len(rows)} 条，全部离线重算（不调用 Dify API）")

    final, improved, regressed, api_errors = recalc(rows, load_script_actions())
    report(rows, final, improved, regressed, api_errors)

    if args.dry_run:
        log("\n--dry-run：未写入文件。")
        return 0

    saved = save(XLSX_PATH, final)
    log(f"\n结果已写入：{saved}")
    if saved != XLSX_PATH:
        log("  原文件被占用，修正后的结果另存为 results_recalc.xlsx；"
            "关闭 Excel 后可改用 --dry-run 之外的方式覆盖原文件。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

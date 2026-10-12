#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
早稲田大学「人を対象とする研究に関する倫理審査申請書」(様式1〜4一体型 Ver2.1)
記載内容の事前チェックスクリプト
"""

import re
import sys
import openpyxl

EXPECTED_VERSION = "Ver.2.1"
EXPECTED_SHEETS = ["様式１", "様式２", "様式３", "様式４"]
VERSION_CELL = "V2"


def check_template_version(wb):
    missing_sheets = [s for s in EXPECTED_SHEETS if s not in wb.sheetnames]
    if missing_sheets:
        return False, (
            f"想定しているシート({'・'.join(missing_sheets)})が見つかりません。"
            "公式テンプレート(様式1〜4一体型)のファイルかどうか確認してください。"
        )

    versions = {}
    for sheet_name in EXPECTED_SHEETS:
        ws = wb[sheet_name]
        v = ws[VERSION_CELL].value
        versions[sheet_name] = v

    mismatched = {
        name: v for name, v in versions.items()
        if str(v).strip() != EXPECTED_VERSION
    }
    if mismatched:
        detail = "、".join(f"{name}:{v}" for name, v in mismatched.items())
        return False, (
            f"テンプレートのバージョン表記が、このスクリプトが動作確認済みの"
            f"バージョン「{EXPECTED_VERSION}」と一致しません({detail})。\n"
            "  大学側でテンプレートが改訂されている可能性があります。改訂されている場合、"
            "行・列の位置がずれていて、このスクリプトのチェック結果が正しく出ない"
            "おそれがあります。\n"
            "  最新の公式テンプレートと見比べて、セル位置に変更がないか確認したうえで、"
            "問題なければ --force オプションを付けて再実行してください。"
        )
    return True, f"テンプレートバージョン確認OK({EXPECTED_VERSION})"


def cell(ws, coord):
    v = ws[coord].value
    return v if v is not None else ""


def is_checked(v):
    return isinstance(v, str) and v.strip() == "■"


def has_text(v):
    return isinstance(v, str) and v.strip() not in ("", "□", "■")


def run_checks(wb):
    ws1 = wb["様式１"]
    ws2 = wb["様式２"]
    ws3 = wb["様式３"]

    results = {"問題あり": [], "要確認": [], "問題なし": []}

    def add(level, title, detail):
        results[level].append((title, detail))

    pi_affil = cell(ws1, "Q11")
    if any(x in pi_affil for x in ("学部", "研究室", "研究科")):
        add("要確認", "様式1 所属欄(研究責任者)",
            f"教員の所属欄には学術院名のみを記載する決まりです。現在の値:「{pi_affil}」。"
            "学部名・研究科名・研究室名が混じっていないか確認してください。")
    else:
        add("問題なし", "様式1 所属欄(研究責任者)", f"「{pi_affil}」")

    rep_affil = cell(ws1, "Q21")
    if any(x in rep_affil for x in ("専攻", "研究室")):
        add("要確認", "様式1 所属欄(研究実施代表者)",
            f"大学院生の所属欄には研究科名のみを記載する決まりです。現在の値:「{rep_affil}」。"
            "専攻名・研究室名が入っていないか確認してください。")
    else:
        add("問題なし", "様式1 所属欄(研究実施代表者)", f"「{rep_affil}」")

    rep_qual = cell(ws1, "R24")
    if ("課程" in rep_qual) and any(x in rep_qual for x in ("年", "回生")):
        add("問題なし", "様式1 資格欄(研究実施代表者)", f"「{rep_qual}」")
    else:
        add("要確認", "様式1 資格欄(研究実施代表者)",
            f"大学院生の場合、博士・修士課程の区分と学年の両方が必要です。現在の値:「{rep_qual}」。")

    start_asap = is_checked(cell(ws2, "D300"))
    start_fixed = is_checked(cell(ws2, "D301"))
    end_year = cell(ws2, "I303")
    end_month = cell(ws2, "L303")
    if start_fixed:
        add("要確認", "様式2 項目10(研究実施期間)",
            "開始日を具体的な日付で指定しています。終了日が開始日から5年以内かどうかを"
            "手動で確認してください(このスクリプトでは開始日の具体的な年月まで読み取っていません)。")
    elif start_asap and isinstance(end_year, int):
        add("問題なし", "様式2 項目10(研究実施期間)",
            f"開始は「承認され次第」、終了は{end_year}年{end_month}月。"
            "承認予定時期から5年以内に収まっているか、念のため目視でも確認してください。")
    else:
        add("要確認", "様式2 項目10(研究実施期間)", "開始・終了日の記載状況を確認できませんでした。")

    a1 = is_checked(cell(ws2, "C265"))
    b1 = is_checked(cell(ws2, "C268"))
    db_register = is_checked(cell(ws2, "B330"))
    if a1 and b1 and not db_register:
        add("問題あり", "様式2 項目12(公開データベース登録)",
            "研究の区分がA①かつB①(生命科学・医学系研究かつ介入あり)のため、"
            "公開データベースへの登録が必須ですが、「登録する」にチェックがありません。")
    else:
        add("問題なし", "様式2 項目12(公開データベース登録)",
            "登録必須の条件(A①かつB①)には該当しないか、条件に対して記載が一致しています。")

    no_company = is_checked(cell(ws2, "B352"))
    econ_none = is_checked(cell(ws2, "A366"))
    econ_none_reason = has_text(cell(ws2, "B368"))
    econ_yes = is_checked(cell(ws2, "A372"))
    econ_yes_detail = has_text(cell(ws2, "B374"))
    if no_company and (econ_none or econ_yes or econ_none_reason or econ_yes_detail):
        add("問題あり", "様式2 項目14/15(企業等との関わり)",
            "項目14で「①企業等との関わりはない」を選択した場合、項目15への回答は不要ですが、"
            "項目15に記載があります。削除してください。")
    elif (not no_company) and not (econ_none or econ_yes):
        add("問題あり", "様式2 項目14/15(企業等との関わり)",
            "項目14で企業等との関わりがある選択肢を選んでいる場合、項目15(経済的利益関係の有無)"
            "への回答が必須ですが、未記載です。")
    else:
        add("問題なし", "様式2 項目14/15(企業等との関わり)", "整合しています。")

    recording = is_checked(cell(ws2, "B492"))
    other_info_checked = is_checked(cell(ws2, "B745"))
    other_info_text = cell(ws2, "F746")
    if recording and not (other_info_checked and "録画" in str(other_info_text) or "録音" in str(other_info_text)):
        add("問題あり", "様式2 項目19/33(録画・録音の記録)",
            "項目19で音声・画像等の記録「②あり」を選択している場合、記録内容が個人を特定し得る"
            "情報であれば、項目33「⑤その他個人を特定し得る情報」欄に録画・録音データについての"
            "記載が必要です。")
    else:
        add("問題なし", "様式2 項目19/33(録画・録音の記録)",
            "録画・録音「あり」を選択していないか、選択している場合は項目33に対応する記載があります。")

    keep_years = cell(ws2, "F504")
    keep_months = cell(ws2, "I504")
    retention_note = cell(ws2, "L504")

    def format_retention_until(end_year, end_month, keep_years, keep_months):
        """研究終了予定年月 + 保管年数・月数 から、保管期限の絶対年月(文字列)を計算する。
        終了予定が具体的な年月で入力されていない場合はNoneを返す。"""
        if not (isinstance(end_year, int) and isinstance(end_month, int)):
            return None
        y = keep_years if isinstance(keep_years, (int, float)) else 0
        m = keep_months if isinstance(keep_months, (int, float)) else 0
        total = end_year * 12 + (end_month - 1) + int(y) * 12 + int(m)
        return f"{total // 12}年{total % 12 + 1}月"

    def parse_absolute_note(text):
        """保管期間欄(L504)の注記文字列から、絶対年月(年, 月)を抜き出す。
        「○○○○年○月」のパターンが見つからなければNoneを返す。"""
        m = re.search(r"(\d{3,4})年(\d{1,2})月", str(text))
        if not m:
            return None
        return int(m.group(1)), int(m.group(2))

    noted = parse_absolute_note(retention_note)
    has_absolute_note = noted is not None

    # 記載されている絶対年月が、「研究終了予定 + 5年」の最低ラインを実際に満たしているかどうか。
    # end_year/end_monthが具体的な年月で入力されていない場合は判定できないためNone(要確認扱い)。
    meets_minimum = None
    if noted and isinstance(end_year, int) and isinstance(end_month, int):
        noted_total = noted[0] * 12 + (noted[1] - 1)
        required_total = end_year * 12 + (end_month - 1) + 5 * 12
        meets_minimum = noted_total >= required_total

    if isinstance(keep_years, (int, float)) and keep_years < 5:
        add("要確認", "様式2 項目20(データの保管期間)",
            f"研究終了後の保管年数が{keep_years}年になっています。5年以上を推奨事項としています。")
    elif not has_absolute_note:
        # 「研究終了後○年○月まで保管する」という相対表記自体は(年金の被保険者期間表記などと同様)
        # 日本語として誤りではないが、委員会からの指摘は「絶対年月を記載すること」そのものなので、
        # 相対表記のみでは指摘事項を解消したことにならない。そのため「要確認」では流さず
        # 「問題あり」として扱う。
        until_text = format_retention_until(end_year, end_month, keep_years, keep_months)
        detail = (f"保管期間欄が「研究終了後{keep_years}年{keep_months}月まで保管する」という"
                   "相対表記のみで記載されており、絶対年月の記載がありません。")
        if until_text:
            detail += f" 研究終了予定({end_year}年{end_month}月)から数えると、{until_text}までに相当します。"
        add("問題あり", "様式2 項目20(データの保管期間)",
            detail + "委員会の指摘事項は絶対年月の記載を求めているため、相対表記だけでは"
                     "解消されません。様式2の保管期間欄に絶対年月(例:「(○○○○年○月まで)」)を"
                     "赤字・下線で追記してください。")
    elif meets_minimum is False:
        add("問題あり", "様式2 項目20(データの保管期間)",
            f"保管期間欄に絶対年月(「{retention_note}」)が記載されていますが、研究終了予定"
            f"({end_year}年{end_month}月)から5年以上先になっていない可能性があります。"
            "記載内容を確認し、研究終了後5年以上先の年月に修正してください。")
    elif meets_minimum is True:
        add("問題なし", "様式2 項目20(データの保管期間)",
            f"保管期間欄に絶対年月(「{retention_note}」)が記載されており、研究終了予定"
            f"({end_year}年{end_month}月)から5年以上先になっていることを確認しました。")
    else:
        # has_absolute_note=True だが、研究終了予定年月が具体的でないなどの理由でmeets_minimumを
        # 計算できなかった場合。絶対年月の記載自体はあるため、5年以上かどうかは目視で確認する。
        add("要確認", "様式2 項目20(データの保管期間)",
            f"保管期間欄に絶対年月(「{retention_note}」)が記載されています。研究終了予定年月が"
            "具体的に読み取れなかったため、5年以上先になっているかは目視で確認してください。")

    youryou_none = is_checked(cell(ws2, "L601"))
    item26_1 = is_checked(cell(ws2, "B612"))
    if youryou_none and item26_1:
        add("問題あり", "様式2 項目25/26(募集要領)",
            "募集要領を添付していないにもかかわらず、項目26①「募集要領に明記している」に"
            "チェックが入っています。募集要領を使わない場合は、このチェックを外し、"
            "様式3(説明文書)に同内容を明記する形にしてください。")
    else:
        add("問題なし", "様式2 項目25/26(募集要領)", "整合しています。")

    written_consent = is_checked(cell(ws2, "B643"))
    form4_attached = is_checked(cell(ws1, "A51"))
    name_checked = is_checked(cell(ws2, "B733"))
    if written_consent and form4_attached and not name_checked:
        add("要確認", "様式2 項目33(収集する個人情報 - 氏名)",
            "様式4(参加への同意書)を添付し、文書によるインフォームド・コンセントを"
            "得る設計になっていますが、項目33「収集する個人情報」の「①氏名」にチェックが"
            "ありません。参加者は同意書に署名するため、通常は①氏名にもチェックと利用目的の"
            "記載が必要です。\n"
            "      ※ただし、氏名を研究データとして保存・利用せず、同意確認のためだけに"
            "事務的に収集・保管する場合は、項目33セルコメントの例外規定により記載不要です"
            "(その場合は「問題なし」扱いで構いません)。")
    else:
        add("問題なし", "様式2 項目33(収集する個人情報 - 氏名)", "整合しています。")

    no_personal_info = is_checked(cell(ws2, "B752"))
    anonymize = is_checked(cell(ws2, "B755"))
    if no_personal_info and anonymize:
        add("問題あり", "様式2 項目33/34(個人情報を収集しない×匿名化)",
            "項目33で「⑥個人情報を収集しない」を選択しているにもかかわらず、項目34で"
            "「①匿名化する」にチェックが入っています。個人情報がなければ匿名化の対象も"
            "存在しないため、矛盾しています。")
    else:
        add("問題なし", "様式2 項目33/34(個人情報を収集しない×匿名化)", "整合しています。")

    pi_extension = cell(ws1, "R12")
    contact_value = ""
    for r in range(221, 232):
        for c in range(1, 20):
            v = ws3.cell(row=r, column=c).value
            if v:
                contact_value += str(v)
    if pi_extension and pi_extension not in contact_value:
        add("要確認", "様式3 項目7(問い合わせ先)",
            f"研究責任者の内線({pi_extension})が、説明文書の問い合わせ先欄に記載されて"
            "いない可能性があります。研究実施代表者の連絡先だけでなく、研究責任者の"
            "研究室の外線電話番号も記載する決まりです。実際のセル内容を目視で確認してください。")
    else:
        add("問題なし", "様式3 項目7(問い合わせ先)", "研究責任者の内線番号が含まれています。")

    return results


def print_report(results, path):
    print("=" * 70)
    print(f"倫理審査申請書 事前チェック結果:{path}")
    print("=" * 70)
    for level in ("問題あり", "要確認", "問題なし"):
        print(f"\n【{level}】 {len(results[level])}件")
        for title, detail in results[level]:
            print(f"  - {title}")
            print(f"      {detail}")
    print("\n" + "=" * 70)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv[1:]

    if len(args) != 1:
        print("使い方: python3 ethics_form_checker.py 対象ファイル.xlsx [--force]")
        sys.exit(1)
    path = args[0]

    wb = openpyxl.load_workbook(path, data_only=True)

    version_ok, version_message = check_template_version(wb)
    print("-" * 70)
    print("テンプレートバージョン確認")
    print("-" * 70)
    print(version_message)
    if not version_ok:
        if force:
            print("\n※--force が指定されているため、警告のまま内容チェックを続行します。")
        else:
            print("\n内容チェックは行わずに終了します。上記の点をご確認ください。")
            sys.exit(1)
    print()

    results = run_checks(wb)
    print_report(results, path)


if __name__ == "__main__":
    main()
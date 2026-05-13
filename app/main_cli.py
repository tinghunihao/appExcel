from pathlib import Path

from logic import (
    MatchColumns,
    complete_address_workbook,
    complete_anhui_address,
    extract_and_merge_workbook,
    match_order_workbooks,
)


def ask(prompt: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip()
    return value or default


def ask_path(prompt: str, default: str = "") -> Path:
    value = ask(prompt, default).strip('"')
    return Path(value)


def pause() -> None:
    input("\n按回车键继续...")


def run_address_preview() -> None:
    address = ask("请输入单条地址", "阜南县朱寨镇刘锋手机大卖场")
    completed, status = complete_anhui_address(address)
    print("\n补全结果：")
    print(completed)
    print(f"状态：{status}")


def run_address_batch() -> None:
    input_path = ask_path("请输入地址 Excel 路径")
    output_path = ask_path("请输入输出 Excel 路径", str(input_path.with_name(input_path.stem + "_地址补全.xlsx")))
    address_column = ask("请输入地址列名", "门店")
    stats = complete_address_workbook(input_path, output_path, address_column)
    print(f"\n处理完成：{output_path}")
    print(f"总行数={stats['rows']}，已补全={stats['changed']}，未识别={stats['unresolved']}")


def run_extract_merge() -> None:
    raw_path = ask_path("请输入原始表 Excel 路径")
    master_path = ask_path("请输入机型主数据表 Excel 路径")
    output_path = ask_path("请输入输出 Excel 路径", str(raw_path.with_name(raw_path.stem + "_字段提取属性匹配.xlsx")))
    selected_columns = ask(
        "请输入提取列，支持列序号或列名，逗号分隔",
        "客户姓名,手机号,产品名称/规格,69国标码,客户实付价,补贴金额,s/n码,Imei-1号,Imei-2号,核销门店,还货地址,联系人,电话",
    )
    rename_spec = ask("请输入列名修改，格式如 1=订单号,69国标码=商品条码，可留空", "")
    raw_code_header = ask("请输入原始表69码列名", "69国标码")
    master_code_header = ask("请输入主数据69码列名", "69码")
    stats = extract_and_merge_workbook(
        raw_path,
        master_path,
        output_path,
        selected_columns,
        rename_spec,
        raw_code_header,
        master_code_header,
    )
    print(f"\n处理完成：{output_path}")
    print(f"总行数={stats['rows']}，匹配成功={stats['matched']}，未找到={stats['not_found']}，提取列数={stats['columns']}")


def run_ab_match() -> None:
    a_path = ask_path("请输入 A表 Excel 路径")
    b_path = ask_path("请输入 B表 Excel 路径")
    output_path = ask_path("请输入输出 Excel 路径", str(a_path.with_name(a_path.stem + "_已匹配发货单号.xlsx")))
    columns = MatchColumns(
        ask("A表电话列", "电话"),
        ask("A表姓名列", "联系人"),
        ask("A表厂家机型列", "厂家机型"),
        ask("A表模糊机型列", "系统机型"),
        ask("B表电话列", "联系方式"),
        ask("B表姓名列", "收货人"),
        ask("B表机型列", "机型版本系列（填写产品匹配里面的格式）"),
        ask("B表发货单号列", "发货通知单号"),
    )
    stats = match_order_workbooks(a_path, b_path, output_path, columns)
    print(f"\n处理完成：{output_path}")
    print(f"总行数={stats['rows']}，匹配成功={stats['matched']}，未找到={stats['not_found']}，多候选={stats['ambiguous']}")


def main() -> None:
    while True:
        print("\n门店 Excel 智能处理工具 x86 兼容版")
        print("1. 单条地址补全测试")
        print("2. 批量地址行政区划补全")
        print("3. 字段提取与机型属性匹配")
        print("4. A/B 表发货单号匹配")
        print("0. 退出")
        choice = ask("请选择功能", "0")
        try:
            if choice == "1":
                run_address_preview()
                pause()
            elif choice == "2":
                run_address_batch()
                pause()
            elif choice == "3":
                run_extract_merge()
                pause()
            elif choice == "4":
                run_ab_match()
                pause()
            elif choice == "0":
                break
            else:
                print("无效选择，请重新输入。")
        except Exception as exc:
            print("\n处理失败：")
            print(exc)
            pause()


if __name__ == "__main__":
    main()

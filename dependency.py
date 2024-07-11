from __future__ import annotations
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
import re
from typing import Set, List, TypeVar
import networkx as nx
import matplotlib.pyplot as plt
from pyvis.network import Network
from openpyxl.utils.cell import rows_from_range

T = TypeVar("T")

ADDRESS_PATTERN = re.compile(r"\b[A-Z]+[1-9][0-9]*(?::[A-Z]+[1-9][0-9]*)?\b")


# セルの数式から依存するセルアドレス一覧を取得する関数
def extract_cells_from_formula(formula) -> List[str]:
    addresses = []
    for match in ADDRESS_PATTERN.findall(formula):
        print(f"match: {match}")
        if ":" in match:
            # 範囲参照
            print(f"range match: {match}")
            addresses.extend([cell for pair in rows_from_range(match) for cell in pair])
        else:
            addresses.append(match)

    return list(set(addresses))


# セルの依存関係を表すクラス
class DependencyTree:
    def __init__(self, sheet_name: str, address: str):
        self.sheet_name: str = sheet_name
        self.address: str = address
        self._dependencies: Set[DependencyTree] = set()
        self.value = None

    def add_dependency(self, cell: "DependencyTree") -> None:
        self._dependencies.add(cell)

    def add_dependencies(self, cells: Set["DependencyTree"]) -> None:
        self._dependencies.update(cells)

    def dependencies(self) -> Set["DependencyTree"]:
        return self._dependencies


# 配列からuniqueな値のみ取り出す関数
def unique_elements_preserve_order(input_list: List[T]) -> List[T]:
    return list(dict.fromkeys(input_list))


# セルアドレスをrootとして依存関係を計算
def get_value_or_function(workbook: "Workbook", root_cell_address: str) -> (str, str):
    sheet: Worksheet = workbook.active
    cell = sheet[root_cell_address]
    if cell.data_type == "f":
        # Cell is function cell
        formula = cell.value
        parent_addresses = extract_cells_from_formula(formula)
        parent_addresses = unique_elements_preserve_order(parent_addresses)
        cur = DependencyTree(sheet.title, root_cell_address)
        for parent_address in parent_addresses:
            parent = get_value_or_function(workbook, parent_address)
            cur.add_dependency(parent)
        return cur
    # Cell is number cell
    return DependencyTree(sheet.title, root_cell_address)


# Dependency Treeを走査して
def dfs(cur: "DependencyTree", n: int = 0):
    print(" " * n + cur.address)
    if n == 0:
        for item in cur.dependencies():
            print(item.address)
    for dependency in cur.dependencies():
        dfs(dependency, n + 1)


def add_dependency_to_graph(node: "DependencyTree", G):
    edges = [(parent.address, node.address) for parent in node.dependencies()]
    G.add_edges_from(edges)
    for parent in node.dependencies():
        add_dependency_to_graph(parent, G)


def plot_dependency(root: "DependencyTree") -> None:
    # 有向グラフを作成
    G = nx.DiGraph()
    add_dependency_to_graph(root, G)
    # nx.draw(G, with_labels=True)
    # plt.show()
    net = Network(directed=True)
    net.from_nx(G)
    net.show("output/dependency_tree.html", notebook=False)


def main():
    book_name = "testcases/" + input("Excel book name:")
    root_cell = input("Target cell address:")
    workbook: Workbook = load_workbook(book_name, data_only=False)
    root = get_value_or_function(workbook, root_cell)
    dfs(root)
    plot_dependency(root)


if __name__ == "__main__":
    main()

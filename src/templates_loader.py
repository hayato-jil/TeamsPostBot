# -*- coding: utf-8 -*-
import os

def _candidate_paths(name: str):
    here = os.path.dirname(os.path.abspath(__file__))  # このファイルがある場所
    cwd = os.getcwd()                                   # 実行時のカレント
    return [
        os.path.join(cwd, name),
        os.path.join(here, name),
        os.path.join(cwd, "templates", name),
        os.path.join(here, "templates", name),
    ]

def load_template(name: str) -> str:
    """
    name で指定したテンプレートを複数候補パスから探索して読み込む。
    UTF-8 で開く。見つからなければ探索パスと cwd を含めて分かりやすく例外。
    """
    tried = []
    for path in _candidate_paths(name):
        tried.append(path)
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
    raise FileNotFoundError(
        "template not found: {n}\n  cwd: {cwd}\n  searched:\n    - {paths}".format(
            n=name, cwd=os.getcwd(), paths="\n    - ".join(tried)
        )
    )

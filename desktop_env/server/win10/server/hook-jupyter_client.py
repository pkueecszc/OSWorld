from PyInstaller.utils.hooks import collect_submodules, copy_metadata

# 强制 PyInstaller 收集 jupyter_client 下的所有子模块
hiddenimports = collect_submodules('jupyter_client')

# 强制 PyInstaller 收集 jupyter_client 的所有元数据，这对于其内部的动态发现机制至关重要
datas = copy_metadata('jupyter_client')
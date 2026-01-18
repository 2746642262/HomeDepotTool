import os
import sys
import json
import pandas as pd
import difflib
import xml.etree.ElementTree as ET
import platform  # 用于检测系统
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QTreeWidget, QTreeWidgetItem, QSpinBox, QGroupBox, 
                             QMessageBox, QLineEdit, QMenu, QCheckBox, QDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView, QInputDialog,
                             QStyle, QProgressDialog) 
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QAction, QColor, QBrush, QFont, QIcon

# ================= 修复 1: 解除递归限制 (防止闪退) =================
sys.setrecursionlimit(20000) 
# ================================================================

# ================= 配置 =================
DEFAULT_OPML_FILE = "Homedepot 后台类目路径.opml" 
ROLE_IS_FOLDER = Qt.ItemDataRole.UserRole + 1 

# 新版颜色池：高区分度
COLOR_PALETTE = [
    "#FF9999", "#99CCFF", "#99FF99", "#FFE066", "#CC99FF",
    "#FFB366", "#66FFFF", "#FF99CC", "#CCFF33", "#DDDDDD"
]
# =======================================

class MatchReviewDialog(QDialog):
    """审核对话框"""
    def __init__(self, fuzzy_matches, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"📷 智能路径修复 - 共 {len(fuzzy_matches)} 条待确认")
        self.resize(1100, 750) 
        self.result_list = [] 
        
        layout = QVBoxLayout(self)
        lbl = QLabel("请对比路径差异：(上方橙色为识别结果，下方绿色为标准结果)")
        
        # 字体适配
        font = QFont()
        font.setBold(True)
        if platform.system() == "Windows":
            font.setFamily("Microsoft YaHei UI")
            font.setPointSize(10)
        else:
            font.setPointSize(12) 
        
        lbl.setFont(font)
        lbl.setStyleSheet("color: #E65100; margin-bottom: 10px; background-color: transparent;")
        layout.addWidget(lbl)
        
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["选择", "错误类型", "路径对比 (上:OCR / 下:CSV)", "应用编码"])
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0, 60) 
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch) 
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.verticalHeader().setDefaultSectionSize(45)
        
        self.fuzzy_matches = fuzzy_matches
        self.table.setRowCount(len(fuzzy_matches) * 2)
        
        # 字体设置
        font_mono = QFont("Menlo" if platform.system() == "Darwin" else "Consolas", 10) 
        font_bold = QFont()
        font_bold.setBold(True)
        if platform.system() == "Windows": font_bold.setFamily("Microsoft YaHei UI")

        for i, match in enumerate(fuzzy_matches):
            row_top = i * 2
            row_bottom = i * 2 + 1
            
            item_check = QTableWidgetItem()
            item_check.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            item_check.setCheckState(Qt.CheckState.Checked)
            self.table.setItem(row_top, 0, item_check)
            self.table.setSpan(row_top, 0, 2, 1) 
            
            type_text = match['match_type']
            item_type = QTableWidgetItem(type_text)
            item_type.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_type.setFont(font_bold)
            if "截断" in type_text:
                item_type.setForeground(QBrush(QColor("purple")))
            self.table.setItem(row_top, 1, item_type)
            self.table.setSpan(row_top, 1, 2, 1)
            
            item_tree = QTableWidgetItem(f"🔴 识别: {match['full_path']}")
            item_tree.setBackground(QColor("#FFF3E0")) 
            item_tree.setFont(font_mono)
            item_tree.setToolTip(match['full_path'])
            self.table.setItem(row_top, 2, item_tree)
            
            item_csv = QTableWidgetItem(f"🟢 标准: {match['csv_path']}")
            item_csv.setBackground(QColor("#E8F5E9")) 
            item_csv.setFont(font_mono)
            item_csv.setToolTip(match['csv_path'])
            self.table.setItem(row_bottom, 2, item_csv)
            
            item_code = QTableWidgetItem(match['code'])
            item_code.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_code.setForeground(QBrush(QColor("blue")))
            item_code.setFont(font_bold)
            self.table.setItem(row_top, 3, item_code)
            self.table.setSpan(row_top, 3, 2, 1)
            
        layout.addWidget(self.table)
        
        btn_box = QHBoxLayout()
        btn_ok = QPushButton("✅ 确认并修复")
        btn_ok.setMinimumHeight(45)
        btn_ok.setFont(font_bold)
        btn_ok.clicked.connect(self.accept_selection)
        
        btn_cancel = QPushButton("全部放弃")
        btn_cancel.setMinimumHeight(45)
        btn_cancel.clicked.connect(self.reject)
        
        btn_box.addStretch()
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_ok)
        layout.addLayout(btn_box)
        
    def accept_selection(self):
        for i in range(0, self.table.rowCount(), 2):
            if self.table.item(i, 0).checkState() == Qt.CheckState.Checked:
                self.result_list.append(self.fuzzy_matches[i // 2])
        self.accept()

class CategoryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Homedepot 智能工作台 (Universal Optimized)")
        self.resize(1400, 900)
        self.settings = QSettings("LiKaixuan_Studio", "CategoryTool_Universal")
        self.current_project_path = None
        self.is_dirty = False
        self.prefix_color_map = {} 
        self.next_color_index = 0
        self.load_counter = 0  # 计数器
        
        # ================= 修复 2: 预加载图标 (极大提升速度) =================
        # 提前获取图标，避免在循环中重复请求系统资源
        self.icon_folder = self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)
        self.icon_file = self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)
        # ===================================================================

        # --- 系统字体自动适配 ---
        if platform.system() == "Windows":
            self.setFont(QFont("Microsoft YaHei UI", 10))
        else:
            self.setFont(QFont(".AppleSystemUIFont", 12)) 
            
        self.init_ui()
        self.tree_widget.itemCollapsed.connect(self.on_item_collapsed_recursive)
        self.tree_widget.itemClicked.connect(self.on_item_clicked)
        self.tree_widget.itemChanged.connect(self.on_item_changed)
        self.load_last_session()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 顶部
        project_group = QGroupBox("📁 项目协同")
        project_layout = QHBoxLayout()
        self.lbl_status = QLabel("状态: 等待加载")
        btn_new = QPushButton("📄 新建")
        btn_new.clicked.connect(self.new_project_from_opml)
        btn_open = QPushButton("📂 打开")
        btn_open.clicked.connect(self.open_project_manual)
        btn_save = QPushButton("💾 保存")
        btn_save.clicked.connect(self.save_project)
        btn_reload = QPushButton("🔄 刷新")
        btn_reload.clicked.connect(self.reload_project)
        project_layout.addWidget(self.lbl_status)
        project_layout.addStretch()
        project_layout.addWidget(btn_new)
        project_layout.addWidget(btn_open)
        project_layout.addWidget(btn_save)
        project_layout.addWidget(btn_reload)
        project_group.setLayout(project_layout)
        main_layout.addWidget(project_group)

        # 操作区
        op_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 搜索...")
        self.search_input.returnPressed.connect(self.perform_fuzzy_search)
        self.btn_csv = QPushButton("导入CSV更新")
        self.btn_csv.clicked.connect(self.load_csv_and_update)
        self.spin_start = QSpinBox()
        self.spin_start.setRange(1, 999999)
        self.spin_start.setValue(2)
        self.spin_start.setPrefix("Start: ")
        self.spin_end = QSpinBox()
        self.spin_end.setRange(1, 999999)
        self.spin_end.setValue(1000)
        self.spin_end.setPrefix("End: ")
        self.btn_export = QPushButton("导出MD")
        self.btn_export.clicked.connect(self.export_markdown)

        op_layout.addWidget(self.search_input)
        op_layout.addWidget(self.btn_csv)
        op_layout.addWidget(self.spin_start)
        op_layout.addWidget(self.spin_end)
        op_layout.addWidget(self.btn_export)
        main_layout.addLayout(op_layout)

        # 底部目录树
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabels(["类目结构", "编码 (Code)", "备注", "收藏"])
        self.tree_widget.setColumnWidth(0, 500)
        self.tree_widget.setColumnWidth(1, 150)
        self.tree_widget.setColumnWidth(2, 200)
        self.tree_widget.setColumnWidth(3, 50)
        
        self.tree_widget.setStyleSheet("""
            QTreeWidget::item:selected {
                background-color: black;
                color: white;
            }
        """)
        
        self.tree_widget.setAlternatingRowColors(False)
        self.tree_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree_widget.customContextMenuRequested.connect(self.open_context_menu)
        main_layout.addWidget(self.tree_widget)

    # --- 辅助方法 ---
    def set_item_type(self, item, is_folder):
        item.setData(0, ROLE_IS_FOLDER, is_folder)
        # ================= 修复 2的应用: 使用缓存图标 =================
        if is_folder:
            item.setIcon(0, self.icon_folder)
        else:
            item.setIcon(0, self.icon_file)
        # ===========================================================

    # --- 搜索逻辑 ---
    def perform_fuzzy_search(self):
        t = self.search_input.text().strip()
        if not t: return
        
        self.tree_widget.clearSelection()
        self.tree_widget.collapseAll()
        
        from PyQt6.QtWidgets import QTreeWidgetItemIterator
        it = QTreeWidgetItemIterator(self.tree_widget)
        best, b_score = None, 0.0
        
        while it.value():
            item = it.value()
            raw = item.data(0, Qt.ItemDataRole.UserRole)
            if raw is None: raw = "" 
            
            path = self.get_full_path(item)
            s = max(difflib.SequenceMatcher(None, t.lower(), path.lower()).ratio(), 
                    difflib.SequenceMatcher(None, t.lower(), raw.lower()).ratio())
            if s > b_score: 
                b_score = s
                best = item
            it += 1
            
        if best and b_score > 0.6:
            self.tree_widget.setCurrentItem(best)
            self.tree_widget.scrollToItem(best)
            best.setSelected(True)
            p = best.parent()
            while p: 
                p.setExpanded(True)
                p = p.parent()
            self.statusBar().showMessage(f"已定位: {best.text(0)}")
        else:
            QMessageBox.warning(self, "提示", "未找到匹配项")

    # --- 核心性能优化区 ---
    def new_project_from_opml(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 OPML", "", "OPML (*.opml)")
        if path:
            try: 
                self.progress = QProgressDialog("正在解析庞大的目录树，请稍候...", "取消", 0, 0, self)
                self.progress.setWindowModality(Qt.WindowModality.WindowModal)
                self.progress.show()
                self.tree_widget.setUpdatesEnabled(False) 
                self.tree_widget.blockSignals(True)       
                tree = ET.parse(path); root = tree.getroot().find('body')
                self.tree_widget.clear()
                self.load_counter = 0 
                self.populate_tree_from_xml(root, self.tree_widget)
                self.current_project_path = None
                self.update_status("新项目")
                self.progress.close()
            except Exception as e: QMessageBox.critical(self, "Error", str(e))
            finally:
                self.tree_widget.setUpdatesEnabled(True) 
                self.tree_widget.blockSignals(False)     

    def populate_tree_from_xml(self, xml_node, parent):
        self.load_counter += 1
        # 心跳包：防止未响应
        if self.load_counter % 50 == 0:
            QApplication.processEvents()
            if hasattr(self, 'progress') and self.progress.wasCanceled(): return
        
        for child in xml_node:
            t = child.get('text') or child.get('title')
            if t:
                item = QTreeWidgetItem(parent)
                item.setText(0, t); item.setData(0, Qt.ItemDataRole.UserRole, t); item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable); self.set_favorite_state(item, False)
                self.populate_tree_from_xml(child, item)
                self.set_item_type(item, item.childCount() > 0)
                item.setExpanded(False)

    def load_project_from_path(self, path):
        try:
            self.progress = QProgressDialog("正在加载项目...", "取消", 0, 0, self)
            self.progress.show()
            QApplication.processEvents()
            with open(path, 'r', encoding='utf-8') as f: data = json.load(f)
            self.tree_widget.setUpdatesEnabled(False); self.tree_widget.blockSignals(True)
            self.tree_widget.clear(); self.prefix_color_map = {}; self.next_color_index = 0; self.load_counter = 0
            self.deserialize_tree(data, self.tree_widget)
            self.current_project_path = path; self.is_dirty = False; self.update_status(f"协同: {os.path.basename(path)}")
            self.settings.setValue("last_project_path", path); self.progress.close()
        except Exception as e: QMessageBox.warning(self, "Error", str(e))
        finally: self.tree_widget.setUpdatesEnabled(True); self.tree_widget.blockSignals(False)

    def deserialize_tree(self, data_list, parent):
        self.load_counter += 1
        if self.load_counter % 50 == 0: QApplication.processEvents()
        
        for node in data_list:
            item = QTreeWidgetItem(parent)
            name = node.get("name", ""); code = node.get("code", ""); remark = node.get("remark", ""); fav = node.get("fav", False)
            is_folder = node.get("is_folder", len(node.get("children", [])) > 0)
            
            item.setText(0, name); item.setData(0, Qt.ItemDataRole.UserRole, name)
            item.setText(1, code); self.apply_color_by_code(item, code)
            item.setText(2, remark); self.set_favorite_state(item, fav)
            self.set_item_type(item, is_folder)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.deserialize_tree(node.get("children", []), item)
            item.setExpanded(node.get("expanded", False))

    # --- CSV功能完好 ---
    def load_csv_and_update(self):
        csv_path, _ = QFileDialog.getOpenFileName(self, "选择 CSV", "", "CSV Files (*.csv)")
        if not csv_path: return
        self.tree_widget.blockSignals(True)
        self.tree_widget.setUpdatesEnabled(False) # 冻结
        try:
            start_idx = max(0, self.spin_start.value() - 2); end_idx = self.spin_end.value() - 2
            df = pd.read_csv(csv_path); df.columns = [c.strip() for c in df.columns]
            subset = df.loc[start_idx:end_idx]
            csv_rules = {}
            for _, row in subset.iterrows():
                p = str(row.get('类目途径', '')).strip(); c = str(row.get('分类', '')).strip()
                if p and c and c != 'nan':
                    norm = self.normalize(p); csv_rules[norm] = {"code": c, "leaf": p.replace("\\", "/").split("/")[-1], "raw_path": p}
            self.pending_exact = []; self.pending_fuzzy = [] 
            self.scan_matches(self.tree_widget.invisibleRootItem(), [], csv_rules)
            final = self.pending_exact 
            if self.pending_fuzzy:
                self.tree_widget.blockSignals(False)
                dialog = MatchReviewDialog(self.pending_fuzzy, self)
                if dialog.exec(): final.extend(dialog.result_list)
                else: QMessageBox.information(self, "提示", "已取消模糊匹配")
                self.tree_widget.blockSignals(True)
            cnt = 0
            for act in final:
                item, code = act['item'], act['code']
                if item.text(1) != code:
                    item.setText(1, code); self.apply_color_by_code(item, code)
                    p = item.parent()
                    while p: p.setExpanded(True); p = p.parent()
                    cnt += 1
            self.is_dirty = True; QMessageBox.information(self, "完成", f"更新: {cnt}")
        except Exception as e: QMessageBox.critical(self, "错误", str(e)); import traceback; traceback.print_exc()
        finally: 
            self.tree_widget.setUpdatesEnabled(True) # 恢复
            self.tree_widget.blockSignals(False)

    def scan_matches(self, item, stack, csv_rules):
        raw = item.data(0, Qt.ItemDataRole.UserRole); ns = stack + [raw] if raw else []
        if item.childCount() > 0:
            for i in range(item.childCount()): self.scan_matches(item.child(i), ns, csv_rules)
        is_folder = item.data(0, ROLE_IS_FOLDER)
        if is_folder: return 
        if raw:
            norm = self.normalize("/".join(ns))
            if norm in csv_rules: self.pending_exact.append({"item": item, "code": csv_rules[norm]['code'], "tree_name": raw, "type": "exact"})
            else: self.check_ocr(item, raw, norm, csv_rules)

    def check_ocr(self, item, raw, norm, rules):
        keys = list(rules.keys()); m = difflib.get_close_matches(norm, keys, n=1, cutoff=0.75)
        if m:
            cand = rules[m[0]]; t_leaf, c_leaf = raw, cand['leaf']
            if len(t_leaf) > len(c_leaf) + 2: return
            if c_leaf.lower().startswith(t_leaf.lower()): self.add_fz(item, cand, raw, "⚠️ 尾部截断", cand['raw_path'])
            elif difflib.SequenceMatcher(None, t_leaf.lower(), c_leaf[:len(t_leaf)].lower()).ratio() > 0.75: self.add_fz(item, cand, raw, "🔡 OCR 错字", cand['raw_path'])

    def add_fz(self, item, dat, t_name, typ, c_path):
        self.pending_fuzzy.append({"item": item, "code": dat['code'], "tree_name": t_name, "full_path": self.get_full_path(item), "csv_path": c_path, "match_type": typ})

    def on_item_collapsed_recursive(self, item):
        self.tree_widget.itemCollapsed.disconnect(self.on_item_collapsed_recursive); self.recursive_collapse(item); self.tree_widget.itemCollapsed.connect(self.on_item_collapsed_recursive)
    def recursive_collapse(self, item):
        for i in range(item.childCount()): child = item.child(i); child.setExpanded(False); self.recursive_collapse(child)
    
    # --- 其他逻辑 ---
    def apply_color_by_code(self, item, code):
        if not code or code.strip() == "":
            for col in range(4): item.setBackground(col, QBrush(Qt.GlobalColor.transparent))
            return
        prefix = code.split('/')[0] if '/' in code else code
        if prefix not in self.prefix_color_map:
            self.prefix_color_map[prefix] = self.next_color_index
            self.next_color_index = (self.next_color_index + 1) % len(COLOR_PALETTE)
        color_idx = self.prefix_color_map[prefix]
        bg_color = QColor(COLOR_PALETTE[color_idx])
        for col in range(4): item.setBackground(col, QBrush(bg_color))

    def on_item_changed(self, item, column):
        if column == 1:
            self.apply_color_by_code(item, item.text(1))
            self.is_dirty = True
        elif column == 2 or column == 0:
            self.is_dirty = True
            if column == 0: item.setData(0, Qt.ItemDataRole.UserRole, item.text(0))

    def set_favorite_state(self, item, is_fav):
        if is_fav:
            item.setText(3, "★")
            item.setForeground(3, QBrush(QColor("#FFD700")))
            item.setData(3, Qt.ItemDataRole.UserRole, True)
        else:
            item.setText(3, "☆")
            item.setForeground(3, QBrush(QColor("#BDBDBD")))
            item.setData(3, Qt.ItemDataRole.UserRole, False)

    def on_item_clicked(self, item, column):
        if column == 3:
            current_state = item.data(3, Qt.ItemDataRole.UserRole)
            self.set_favorite_state(item, not current_state)
            self.is_dirty = True

    def open_context_menu(self, pos):
        item = self.tree_widget.itemAt(pos)
        menu = QMenu()
        if item:
            ac_rename = QAction("✏️ 重命名", self)
            ac_rename.triggered.connect(lambda: self.action_rename(item))
            menu.addAction(ac_rename)
            ac_del = QAction("🗑️ 删除", self)
            ac_del.triggered.connect(lambda: self.action_delete(item))
            menu.addAction(ac_del)
            
            is_folder = item.data(0, ROLE_IS_FOLDER)
            if is_folder:
                menu.addSeparator()
                ac_add_folder = QAction("📂 新建子目录", self)
                ac_add_folder.triggered.connect(lambda: self.action_add_child(item, True))
                menu.addAction(ac_add_folder)
                ac_add_file = QAction("📄 新建子文件", self)
                ac_add_file.triggered.connect(lambda: self.action_add_child(item, False))
                menu.addAction(ac_add_file)
            
            menu.addSeparator()
            ac_copy = QAction("📋 复制路径", self)
            ac_copy.triggered.connect(lambda: QApplication.clipboard().setText(self.get_full_path(item)))
            menu.addAction(ac_copy)
        else:
            ac_add_root = QAction("➕ 新建顶级类目", self)
            ac_add_root.triggered.connect(lambda: self.action_add_child(self.tree_widget.invisibleRootItem(), True))
            menu.addAction(ac_add_root)
        menu.exec(self.tree_widget.viewport().mapToGlobal(pos))

    def action_rename(self, item):
        text, ok = QInputDialog.getText(self, "重命名", "新名称:", text=item.text(0))
        if ok and text: item.setText(0, text)
    
    def action_delete(self, item):
        reply = QMessageBox.question(self, "确认删除", f"确定要删除 '{item.text(0)}' 吗？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            parent = item.parent() or self.tree_widget.invisibleRootItem()
            parent.removeChild(item)
            self.is_dirty = True

    def action_add_child(self, parent, is_folder):
        type_n = "目录" if is_folder else "文件"
        text, ok = QInputDialog.getText(self, f"新建{type_n}", "名称:")
        if ok and text:
            new = QTreeWidgetItem(parent)
            new.setText(0, text)
            new.setData(0, Qt.ItemDataRole.UserRole, text)
            self.set_item_type(new, is_folder)
            new.setFlags(new.flags() | Qt.ItemFlag.ItemIsEditable)
            self.set_favorite_state(new, False)
            parent.setExpanded(True)
            self.is_dirty = True

    def serialize_tree(self, item):
        children = []
        for i in range(item.childCount()): children.append(self.serialize_tree(item.child(i)))
        return {
            "name": item.data(0, Qt.ItemDataRole.UserRole),
            "code": item.text(1),
            "remark": item.text(2),
            "fav": item.data(3, Qt.ItemDataRole.UserRole),
            "is_folder": item.data(0, ROLE_IS_FOLDER),
            "expanded": item.isExpanded(),
            "children": children
        }

    def open_project_manual(self):
        path, _ = QFileDialog.getOpenFileName(self, "JSON", "", "JSON (*.json)")
        if path: self.load_project_from_path(path)

    def save_project(self):
        if not self.current_project_path: path, _ = QFileDialog.getSaveFileName(self, "Save", "Team.json", "JSON (*.json)"); 
        else: path = self.current_project_path
        if not path: return
        self.current_project_path = path; data = self.serialize_tree(self.tree_widget.invisibleRootItem())
        with open(self.current_project_path, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False, indent=2)
        self.settings.setValue("last_project_path", self.current_project_path); self.is_dirty = False; self.update_status(f"Saved: {os.path.basename(self.current_project_path)}")
    
    def load_last_session(self):
        p = self.settings.value("last_project_path")
        if p and os.path.exists(p): self.load_project_from_path(p)
    
    def reload_project(self):
        if self.current_project_path: self.load_project_from_path(self.current_project_path)
    
    def update_status(self, t): self.lbl_status.setText(f"状态: {t}"); self.lbl_status.setStyleSheet("color: red; font-weight: bold;" if "未保存" in t else "color: green;")
    
    def export_markdown(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export", "Catalog.txt", "Text (*.txt)")
        if path:
            with open(path, 'w', encoding='utf-8') as f: self.write_md(self.tree_widget.invisibleRootItem(), f, -1)
            QMessageBox.information(self, "Success", "Done")
    
    def write_md(self, item, f, level):
        if level >= 0:
            ind = "    " * level; rem = f" # {item.text(2)}" if item.text(2) else ""
            f.write(f"{ind}- {item.text(0)} {item.text(1)}{rem}\n")
        for i in range(item.childCount()): self.write_md(item.child(i), f, level + 1)
    
    def normalize(self, t): return t.lower().replace(" ", "").replace(">", "/").replace("\\", "/")
    
    def get_full_path(self, item):
        p = []; c = item
        while c: p.insert(0, c.data(0, Qt.ItemDataRole.UserRole)); c = c.parent()
        return "/".join(p)

if __name__ == "__main__":
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    
    # 全局强制白天模式
    app.setStyleSheet("""
        QWidget { background-color: #FFFFFF; color: #000000; }
        QTreeWidget, QTableWidget, QListWidget { background-color: #FFFFFF; color: #000000; border: 1px solid #D0D0D0; alternate-background-color: #FFFFFF; }
        QLineEdit, QSpinBox, QTextEdit { background-color: #FFFFFF; color: #000000; border: 1px solid #C0C0C0; border-radius: 3px; padding: 2px; }
        QHeaderView::section { background-color: #F0F0F0; color: #000000; border: 1px solid #D8D8D8; padding: 4px; }
        QPushButton { background-color: #F5F5F5; color: #000000; border: 1px solid #C0C0C0; border-radius: 3px; padding: 5px 10px; }
        QPushButton:hover { background-color: #E0E0E0; }
        QPushButton:pressed { background-color: #D0D0D0; }
        QMenu { background-color: #FFFFFF; color: #000000; border: 1px solid #CCCCCC; }
        QMenu::item:selected { background-color: #0078D7; color: #FFFFFF; }
    """)
    
    window = CategoryApp()
    window.show()
    sys.exit(app.exec())

# powerbb_ui.py — Single-file PySide6 GUI for “PowerBB AI Slide Manager”
# Widgets/Layout constraints honored; hooks to powerbb.py backend.

from PySide6 import QtCore, QtGui, QtWidgets
import json, os, traceback
import powerbb

# Import your backend sitting next to this file
import powerbb  # <-- make sure powerbb.py is in the same folder

SPACING = 8
PADDING = 12
RADIUS = 8
FONT_BASE_PT = 12

SETTINGS_FILE = "powerbb_ai_slide_manager.settings.json"

class SlideManagerWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PowerBB AI Slide Manager")
        self.setMinimumSize(1100, 720)

        central = QtWidgets.QWidget()
        root_v = QtWidgets.QVBoxLayout(central)
        root_v.setContentsMargins(PADDING, PADDING, PADDING, PADDING)
        root_v.setSpacing(SPACING)

        # Header
        header = QtWidgets.QWidget()
        h = QtWidgets.QHBoxLayout(header)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(SPACING)
        self.title_ro = QtWidgets.QLineEdit("PowerBB AI Slide Manager")
        self.title_ro.setReadOnly(True)
        self.title_ro.setObjectName("HeaderTitle")
        h.addWidget(self.title_ro, 1)
        self.btn_build = QtWidgets.QPushButton("Build deck")
        self.btn_tests = QtWidgets.QPushButton("Run tests")
        self.btn_dump = QtWidgets.QPushButton("Dump layouts")
        self.btn_prompt = QtWidgets.QPushButton("Generate prompt")
        for b in (self.btn_build, self.btn_tests, self.btn_dump, self.btn_prompt):
            h.addWidget(b, 0)
        root_v.addWidget(header, 0)

        # Splitter
        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        root_v.addWidget(self.splitter, 1)

        # Left column (fixed ~480)
        self.left_col = QtWidgets.QWidget()
        self.left_col.setObjectName("LeftColumn")
        self.left_col.setMinimumWidth(480)
        self.left_col.setMaximumWidth(480)
        left_v = QtWidgets.QVBoxLayout(self.left_col)
        left_v.setContentsMargins(0, 0, 0, 0)
        left_v.setSpacing(SPACING)

        # Template group
        self.grp_template = QtWidgets.QGroupBox("Template (.pptx)")
        f = QtWidgets.QFormLayout(self.grp_template)
        f.setHorizontalSpacing(SPACING); f.setVerticalSpacing(SPACING)
        self.le_template = QtWidgets.QLineEdit()
        self.btn_pick_template = QtWidgets.QPushButton("Pick…")
        row_t = QtWidgets.QHBoxLayout(); row_t.setSpacing(SPACING)
        row_t.addWidget(self.le_template, 1); row_t.addWidget(self.btn_pick_template, 0)
        f.addRow("Template:", row_t)
        left_v.addWidget(self.grp_template)

        # Content group
        self.grp_content = QtWidgets.QGroupBox("Content (PowerBB JSON)")
        cv = QtWidgets.QVBoxLayout(self.grp_content); cv.setSpacing(SPACING)
        radios_row = QtWidgets.QHBoxLayout(); radios_row.setSpacing(SPACING)
        self.rb_file = QtWidgets.QRadioButton("File")
        self.rb_clip = QtWidgets.QRadioButton("Clipboard")
        self.rb_paste = QtWidgets.QRadioButton("Paste")
        self.rb_file.setChecked(True)
        radios_row.addWidget(self.rb_file); radios_row.addWidget(self.rb_clip); radios_row.addWidget(self.rb_paste)
        radios_row.addStretch(1)
        cv.addLayout(radios_row)
        file_row = QtWidgets.QHBoxLayout(); file_row.setSpacing(SPACING)
        self.le_json_file = QtWidgets.QLineEdit()
        self.btn_browse_json = QtWidgets.QPushButton("Browse…")
        file_row.addWidget(self.le_json_file, 1); file_row.addWidget(self.btn_browse_json, 0)
        cv.addLayout(file_row)
        self.paste_edit = QtWidgets.QPlainTextEdit()
        self.paste_edit.setPlaceholderText("Paste PowerBB JSON here…")
        self.paste_edit.setObjectName("PasteArea")
        cv.addWidget(self.paste_edit, 1)
        left_v.addWidget(self.grp_content, 1)

        # Outputs group
        self.grp_outputs = QtWidgets.QGroupBox("Outputs")
        of = QtWidgets.QFormLayout(self.grp_outputs)
        of.setHorizontalSpacing(SPACING); of.setVerticalSpacing(SPACING)
        self.le_out_deck = QtWidgets.QLineEdit(); self.btn_out_deck = QtWidgets.QPushButton("Browse…")
        r1 = QtWidgets.QHBoxLayout(); r1.setSpacing(SPACING)
        r1.addWidget(self.le_out_deck, 1); r1.addWidget(self.btn_out_deck, 0)
        of.addRow("Deck (.pptx):", r1)
        self.le_out_profile = QtWidgets.QLineEdit(); self.btn_out_profile = QtWidgets.QPushButton("Browse…")
        r2 = QtWidgets.QHBoxLayout(); r2.setSpacing(SPACING)
        r2.addWidget(self.le_out_profile, 1); r2.addWidget(self.btn_out_profile, 0)
        of.addRow("Template profile (.json):", r2)
        self.le_out_prompt = QtWidgets.QLineEdit(); self.btn_out_prompt = QtWidgets.QPushButton("Browse…")
        r3 = QtWidgets.QHBoxLayout(); r3.setSpacing(SPACING)
        r3.addWidget(self.le_out_prompt, 1); r3.addWidget(self.btn_out_prompt, 0)
        of.addRow("Authoring prompt (.txt):", r3)
        left_v.addWidget(self.grp_outputs)

        # Options group
        self.grp_options = QtWidgets.QGroupBox("Options")
        grid = QtWidgets.QGridLayout(self.grp_options)
        grid.setHorizontalSpacing(SPACING); grid.setVerticalSpacing(SPACING)
        self.cb_lenient = QtWidgets.QCheckBox("Lenient JSON cleanup")
        self.cb_open_folder = QtWidgets.QCheckBox("Open output folder after build")
        self.cb_remember = QtWidgets.QCheckBox("Remember settings")
        self.le_verbosity = QtWidgets.QLineEdit(); self.le_verbosity.setPlaceholderText("0, 1, or 2")
        grid.addWidget(self.cb_lenient, 0, 0, 1, 2)
        grid.addWidget(self.cb_open_folder, 1, 0, 1, 2)
        fake_label = QtWidgets.QLineEdit("Verbosity (-v count)"); fake_label.setReadOnly(True); fake_label.setObjectName("FakeLabel")
        grid.addWidget(fake_label, 2, 0); grid.addWidget(self.le_verbosity, 2, 1)
        grid.addWidget(self.cb_remember, 3, 0, 1, 2)
        left_v.addWidget(self.grp_options)

        spacer = QtWidgets.QWidget()
        spacer.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Expanding)
        left_v.addWidget(spacer, 1)

        # Right column
        self.right_col = QtWidgets.QWidget()
        right_v = QtWidgets.QVBoxLayout(self.right_col)
        right_v.setContentsMargins(0, 0, 0, 0); right_v.setSpacing(SPACING)
        self.card_preview = QtWidgets.QGroupBox("AI Prompt Preview"); self.card_preview.setObjectName("Card")
        pv = QtWidgets.QVBoxLayout(self.card_preview); pv.setContentsMargins(PADDING, PADDING, PADDING, PADDING); pv.setSpacing(SPACING)
        self.preview_edit = QtWidgets.QPlainTextEdit(); self.preview_edit.setReadOnly(True)
        pv.addWidget(self.preview_edit, 1)
        right_v.addWidget(self.card_preview, 2)
        self.card_logs = QtWidgets.QGroupBox("Logs"); self.card_logs.setObjectName("Card")
        lv = QtWidgets.QVBoxLayout(self.card_logs); lv.setContentsMargins(PADDING, PADDING, PADDING, PADDING); lv.setSpacing(SPACING)
        self.log_edit = QtWidgets.QPlainTextEdit()
        mono = QtGui.QFontDatabase.systemFont(QtGui.QFontDatabase.SystemFont.FixedFont); mono.setPointSize(FONT_BASE_PT)
        self.log_edit.setFont(mono)
        lv.addWidget(self.log_edit, 1)
        right_v.addWidget(self.card_logs, 3)

        self.splitter.addWidget(self.left_col); self.splitter.addWidget(self.right_col)
        self.splitter.setCollapsible(0, False); self.splitter.setCollapsible(1, False)
        self.splitter.setStretchFactor(0, 0); self.splitter.setStretchFactor(1, 1)

        self.status = QtWidgets.QStatusBar(); self.setStatusBar(self.status); self.status.showMessage("Ready")
        self.setCentralWidget(central)
        self.apply_qss()

        # Connections
        self.btn_pick_template.clicked.connect(self.pick_template)
        self.btn_browse_json.clicked.connect(self.pick_json)
        self.btn_out_deck.clicked.connect(lambda: self.pick_save_path(self.le_out_deck, "PowerPoint (*.pptx)"))
        self.btn_out_profile.clicked.connect(lambda: self.pick_save_path(self.le_out_profile, "JSON (*.json)"))
        self.btn_out_prompt.clicked.connect(lambda: self.pick_save_path(self.le_out_prompt, "Text (*.txt)"))
        self.btn_build.clicked.connect(self.build_deck)
        self.btn_tests.clicked.connect(self.run_tests_stub)
        self.btn_dump.clicked.connect(self.dump_layouts_stub)
        self.btn_prompt.clicked.connect(self.generate_prompt_stub)

        for w in (self.le_template, self.le_json_file, self.le_out_deck, self.le_out_profile,
                  self.le_out_prompt, self.paste_edit, self.le_verbosity):
            if isinstance(w, QtWidgets.QPlainTextEdit):
                w.textChanged.connect(self.validate)
            else:
                w.textChanged.connect(self.validate)
        for rb in (self.rb_file, self.rb_clip, self.rb_paste):
            rb.toggled.connect(self.update_content_mode)
        self.cb_remember.toggled.connect(self.on_remember_toggled)

        self.update_content_mode()
        self.validate()
        self.load_settings_if_opted()

    # ---------- Styling ----------
    def apply_qss(self):
        qss = f"""
        * {{ font-size: {FONT_BASE_PT}pt; }}
        QGroupBox {{
            border: 1px solid #d0d5dd; border-radius: {RADIUS}px;
            margin-top: {PADDING}px; padding: {PADDING - 6}px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin; subcontrol-position: top left;
            padding: 0 {SPACING}px; color: #0f172a; font-weight: 600;
        }}
        #Card {{ background: #fafafa; border: 1px solid #e5e7eb; border-radius: {RADIUS}px; }}
        #HeaderTitle {{ border: none; background: transparent; font-weight: 700; font-size: {FONT_BASE_PT + 2}pt; color: #111827; }}
        QPlainTextEdit {{ border: 1px solid #e5e7eb; border-radius: {RADIUS}px; padding: {PADDING - 6}px; background: #ffffff; }}
        QPushButton {{ padding: 6px 10px; border: 1px solid #d1d5db; border-radius: {RADIUS}px; background: #f3f4f6; }}
        QPushButton:hover {{ background: #e5e7eb; }}
        QPushButton:disabled {{ color: #9ca3af; background: #f9fafb; border-color: #e5e7eb; }}
        #FakeLabel {{ border: none; background: transparent; color: #374151; }}
        """
        self.setStyleSheet(qss)

    # ---------- File pickers ----------
    def pick_template(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select PowerPoint template", "", "PowerPoint (*.pptx)")
        if path:
            self.le_template.setText(path)
            self.status.showMessage(f"Template selected: {path}", 5000)
            self.preview_edit.setPlainText(self.make_preview_text(template=path))
        self.validate()

    def pick_json(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select PowerBB JSON file", "", "JSON (*.json);;All files (*)")
        if path:
            self.le_json_file.setText(path)
            self.status.showMessage(f"JSON selected: {path}", 5000)
        self.validate()

    def pick_save_path(self, line_edit: QtWidgets.QLineEdit, filter_str: str):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Choose output path", "", filter_str)
        if path:
            line_edit.setText(path)
        self.validate()

    # ---------- Modes & validation ----------
    def update_content_mode(self):
        is_file = self.rb_file.isChecked()
        is_paste = self.rb_paste.isChecked()
        self.le_json_file.setEnabled(is_file)
        self.btn_browse_json.setEnabled(is_file)
        self.paste_edit.setEnabled(is_paste)
        self.paste_edit.setVisible(is_paste)
        self.validate()

    def validate(self):
        tmpl_ok = bool(self.le_template.text().strip())
        deck_ok = bool(self.le_out_deck.text().strip())
        if self.rb_file.isChecked():
            content_ok = bool(self.le_json_file.text().strip())
        elif self.rb_paste.isChecked():
            content_ok = bool(self.paste_edit.toPlainText().strip())
        else:
            content_ok = True  # Clipboard will be read at runtime
        self.btn_build.setEnabled(deck_ok and content_ok)
        self.btn_dump.setEnabled(tmpl_ok)
        self.btn_prompt.setEnabled(tmpl_ok)
        self.btn_tests.setEnabled(tmpl_ok)

    # ---------- Backend glue ----------
    def _read_powerbb_json(self):
        if self.rb_file.isChecked():
            p = self.le_json_file.text().strip()
            with open(p, "r", encoding="utf-8") as f:
                raw = f.read()
        elif self.rb_paste.isChecked():
            raw = self.paste_edit.toPlainText()
        else:  # Clipboard
            raw = QtGui.QGuiApplication.clipboard().text()
        if self.cb_lenient.isChecked():
            raw = powerbb.clean_json_lenient(raw)
        return json.loads(raw)

    def append_log(self, text: str):
        self.log_edit.appendPlainText(text)
        cursor = self.log_edit.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        self.log_edit.setTextCursor(cursor)

    def build_deck(self):
        try:
            pb_obj = self._read_powerbb_json()
            out_path = self.le_out_deck.text().strip()
            tmpl = self.le_template.text().strip() or None
            verbosity = self.le_verbosity.text().strip()
            vcount = int(verbosity) if verbosity.isdigit() else 0

            self.append_log("[build] Creating deck…")
            powerbb.create_ppt_from_powerbb(pb_obj, out_path, template_path=tmpl)
            self.append_log(f"[build] Wrote: {out_path}")
            if self.cb_open_folder.isChecked() and out_path:
                folder = QtCore.QFileInfo(out_path).absolutePath()
                self.append_log(f"[build] (info) Would open folder: {folder}")
            self.status.showMessage("Build completed.", 4000)
        except Exception:
            self.append_log(traceback.format_exc())
            self.status.showMessage("Build failed (see Logs).", 6000)

    def run_tests_stub(self):
        try:
            out = self.le_out_deck.text().strip() or None
            tmpl = self.le_template.text().strip() or None
            self.append_log("[tests] Running backend round-trip test…")
            powerbb.test_powerbb_roundtrip(tmp_output_path=out, template_path=tmpl)
            if out:
                self.append_log(f"[tests] Test deck written to: {out}")
            self.append_log("[tests] OK.")
            self.status.showMessage("Tests finished.", 4000)
        except Exception:
            self.append_log(traceback.format_exc())
            self.status.showMessage("Tests failed (see Logs).", 6000)

    def dump_layouts_stub(self):
        try:
            tmpl = self.le_template.text().strip()
            out_json = self.le_out_profile.text().strip() or None
            self.append_log("[dump] Generating template profile…")
            prs = powerbb.Presentation(tmpl) if tmpl else powerbb.Presentation()
            powerbb._dump_layouts(prs, as_json=out_json)
            if out_json:
                self.append_log(f"[dump] Wrote template profile JSON: {out_json}")
            self.append_log("[dump] Done.")
            self.status.showMessage("Dump layouts finished.", 4000)
        except Exception:
            self.append_log(traceback.format_exc())
            self.status.showMessage("Dump failed (see Logs).", 6000)

    def generate_prompt_stub(self):
        try:
            tmpl = self.le_template.text().strip() or None
            self.append_log("[prompt] Building authoring prompt from template inventory…")
            txt = powerbb.generate_powerbb_prompt(tmpl)
            self.preview_edit.setPlainText(txt)
            out = self.le_out_prompt.text().strip()
            if out:
                with open(out, "w", encoding="utf-8") as f:
                    f.write(txt)
                self.append_log(f"[prompt] Wrote prompt to: {out}")
            self.status.showMessage("Prompt generated.", 4000)
        except Exception:
            self.append_log(traceback.format_exc())
            self.status.showMessage("Prompt generation failed (see Logs).", 6000)

    # ---------- Settings ----------
    def settings_path(self) -> str:
        loc = QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.AppConfigLocation)
        if not loc:
            loc = QtCore.QDir.homePath()
        QtCore.QDir().mkpath(loc)
        return QtCore.QDir(loc).filePath(SETTINGS_FILE)

    def on_remember_toggled(self, checked: bool):
        if checked:
            self.save_settings()
        else:
            p = self.settings_path()
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass

    def save_settings(self):
        if not self.cb_remember.isChecked():
            return
        data = {
            "template": self.le_template.text(),
            "content_mode": "file" if self.rb_file.isChecked() else ("clip" if self.rb_clip.isChecked() else "paste"),
            "json_file": self.le_json_file.text(),
            "paste_text": self.paste_edit.toPlainText(),
            "out_deck": self.le_out_deck.text(),
            "out_profile": self.le_out_profile.text(),
            "out_prompt": self.le_out_prompt.text(),
            "lenient": self.cb_lenient.isChecked(),
            "open_folder": self.cb_open_folder.isChecked(),
            "verbosity": self.le_verbosity.text(),
            "remember": self.cb_remember.isChecked(),
        }
        try:
            with open(self.settings_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load_settings_if_opted(self):
        p = self.settings_path()
        if not os.path.exists(p):
            return
        try:
            with open(p, "r", encoding="utf-8") as f:
                obj = json.load(f)
        except Exception:
            return
        remember = bool(obj.get("remember", False))
        self.cb_remember.setChecked(remember)
        if not remember:
            return
        self.le_template.setText(obj.get("template", ""))
        mode = obj.get("content_mode", "file")
        self.rb_file.setChecked(mode == "file")
        self.rb_clip.setChecked(mode == "clip")
        self.rb_paste.setChecked(mode == "paste")
        self.le_json_file.setText(obj.get("json_file", ""))
        self.paste_edit.setPlainText(obj.get("paste_text", ""))
        self.le_out_deck.setText(obj.get("out_deck", ""))
        self.le_out_profile.setText(obj.get("out_profile", ""))
        self.le_out_prompt.setText(obj.get("out_prompt", ""))
        self.cb_lenient.setChecked(bool(obj.get("lenient", False)))
        self.cb_open_folder.setChecked(bool(obj.get("open_folder", False)))
        self.le_verbosity.setText(obj.get("verbosity", "0"))
        self.update_content_mode(); self.validate()

    # ---------- Misc ----------
    def keyPressEvent(self, event: QtGui.QKeyEvent):
        if event.key() == QtCore.Qt.Key.Key_O and (event.modifiers() & QtCore.Qt.KeyboardModifier.ControlModifier):
            self.pick_template(); event.accept(); return
        super().keyPressEvent(event)

    def closeEvent(self, e: QtGui.QCloseEvent):
        self.save_settings()
        super().closeEvent(e)

    def make_preview_text(self, template: str | None) -> str:
        lines = []
        lines.append("ROLE: You generate PowerPoint content as powerbb JSON for a Python builder.")
        lines.append("TASK: Produce ONE valid JSON object; output strict JSON only.")
        if template:
            lines.append(f"TEMPLATE: {template}")
        lines.append("")
        lines.append('{"meta":{"template_path":"<TEMPLATE>.pptx","defaults":{"list_type":"bullet","title_size_pt":40,"body_size_pt":24}},')
        lines.append(' "slides":[{"layout":"Title and Content","title":"Executive Summary — {{client}} ({{year}})","regions":{"left":{"bullets":[{"text":"Point"}]}}}]}')
        return "\n".join(lines)

def main():
    app = QtWidgets.QApplication([])
    w = SlideManagerWindow()
    w.show()
    return app.exec()

if __name__ == "__main__":
    raise SystemExit(main())

import os
import sys
import csv
import json
import sqlite3
import shutil
import hashlib
import zipfile
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

from face_engine import FaceEngine

APP_NAME = 'Фотокаталог'
IMAGE_PATTERNS = ('*.jpg', '*.jpeg', '*.png', '*.bmp', '*.tif', '*.tiff')


def app_data_dir():
    if getattr(sys, 'frozen', False):
        base = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
        root = Path(base) / 'PhotoArchiveCatalog'
    else:
        root = Path(__file__).resolve().parent / 'data'
    root.mkdir(parents=True, exist_ok=True)
    (root / 'photos').mkdir(parents=True, exist_ok=True)
    (root / 'thumbs').mkdir(parents=True, exist_ok=True)
    return root


DATA_DIR = app_data_dir()
DB_PATH = DATA_DIR / 'catalog.db'
PHOTOS_DIR = DATA_DIR / 'photos'
THUMBS_DIR = DATA_DIR / 'thumbs'


class CatalogDB:
    def __init__(self, path):
        self.path = str(path)
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.create_schema()

    def create_schema(self):
        self.conn.executescript('''
        CREATE TABLE IF NOT EXISTS photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            archive_no TEXT NOT NULL DEFAULT '',
            description TEXT NOT NULL DEFAULT '',
            shot_date TEXT NOT NULL DEFAULT '',
            location TEXT NOT NULL DEFAULT '',
            author TEXT NOT NULL DEFAULT '',
            source TEXT NOT NULL DEFAULT '',
            file_path TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS face_index (
            photo_id INTEGER PRIMARY KEY,
            embedding TEXT NOT NULL,
            indexed_at TEXT NOT NULL,
            FOREIGN KEY(photo_id) REFERENCES photos(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_photos_archive_no ON photos(archive_no);
        CREATE INDEX IF NOT EXISTS idx_photos_shot_date ON photos(shot_date);
        CREATE INDEX IF NOT EXISTS idx_photos_author ON photos(author);
        ''')
        self.conn.commit()

    def add(self, record):
        now = datetime.now().isoformat(timespec='seconds')
        cur = self.conn.execute('''INSERT INTO photos
            (archive_no, description, shot_date, location, author, source, file_path, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (record['archive_no'], record['description'], record['shot_date'], record['location'],
             record['author'], record['source'], record['file_path'], now, now))
        self.conn.commit()
        return cur.lastrowid

    def update(self, photo_id, record):
        now = datetime.now().isoformat(timespec='seconds')
        self.conn.execute('''UPDATE photos SET archive_no=?, description=?, shot_date=?, location=?, author=?, source=?,
            file_path=?, updated_at=? WHERE id=?''',
            (record['archive_no'], record['description'], record['shot_date'], record['location'],
             record['author'], record['source'], record['file_path'], now, photo_id))
        self.conn.execute('DELETE FROM face_index WHERE photo_id=?', (photo_id,))
        self.conn.commit()

    def delete(self, photo_id):
        row = self.get(photo_id)
        self.conn.execute('DELETE FROM face_index WHERE photo_id=?', (photo_id,))
        self.conn.execute('DELETE FROM photos WHERE id=?', (photo_id,))
        self.conn.commit()
        return row

    def get(self, photo_id):
        return self.conn.execute('SELECT * FROM photos WHERE id=?', (photo_id,)).fetchone()

    def all(self):
        return self.conn.execute('SELECT * FROM photos ORDER BY id DESC').fetchall()

    def search(self, text=''):
        text = (text or '').strip()
        if not text:
            return self.all()
        tokens = [t for t in text.split() if t]
        fields = ['archive_no', 'description', 'shot_date', 'location', 'author', 'source']
        clauses, params = [], []
        for token in tokens:
            sub = ' OR '.join([f'{f} LIKE ?' for f in fields])
            clauses.append(f'({sub})')
            params.extend([f'%{token}%'] * len(fields))
        sql = 'SELECT * FROM photos WHERE ' + ' AND '.join(clauses) + ' ORDER BY id DESC'
        return self.conn.execute(sql, params).fetchall()

    def set_embedding(self, photo_id, embedding):
        payload = json.dumps([float(x) for x in embedding])
        now = datetime.now().isoformat(timespec='seconds')
        self.conn.execute('INSERT OR REPLACE INTO face_index(photo_id, embedding, indexed_at) VALUES (?, ?, ?)',
                          (photo_id, payload, now))
        self.conn.commit()

    def embeddings(self):
        return self.conn.execute('''SELECT f.photo_id, f.embedding, p.* FROM face_index f
                                    JOIN photos p ON p.id=f.photo_id''').fetchall()

    def close(self):
        self.conn.close()


def copy_into_archive(src_path):
    src = Path(src_path)
    if not src.exists():
        raise FileNotFoundError(str(src))
    digest = hashlib.sha1((str(src.resolve()) + str(src.stat().st_mtime_ns)).encode('utf-8', 'ignore')).hexdigest()[:12]
    suffix = src.suffix.lower() or '.jpg'
    dest = PHOTOS_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{digest}{suffix}"
    shutil.copy2(str(src), str(dest))
    return str(dest)


def make_thumb(path, key, size=(90, 70)):
    if Image is None or not path or not os.path.exists(path):
        return None
    target = THUMBS_DIR / f'{key}.jpg'
    try:
        if not target.exists() or target.stat().st_mtime < Path(path).stat().st_mtime:
            img = Image.open(path)
            img.thumbnail(size)
            bg = Image.new('RGB', size, 'white')
            x = (size[0] - img.width) // 2
            y = (size[1] - img.height) // 2
            if img.mode != 'RGB':
                img = img.convert('RGB')
            bg.paste(img, (x, y))
            bg.save(str(target), 'JPEG', quality=85)
        return str(target)
    except Exception:
        return None


class PhotoCatalogApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry('1380x820')
        self.minsize(1100, 680)
        self.db = CatalogDB(DB_PATH)
        self.face = FaceEngine()
        self.current_id = None
        self.current_image_path = ''
        self.preview_ref = None
        self.thumb_refs = {}
        self.protocol('WM_DELETE_WINDOW', self.on_close)
        self.build_ui()
        self.load_rows()

    def build_ui(self):
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=0)
        file_menu.add_command(label='Массовый импорт фотографий…', command=self.bulk_import_photos)
        file_menu.add_command(label='Импорт Excel / CSV…', command=self.import_table)
        file_menu.add_command(label='Импорт Access…', command=self.import_access)
        file_menu.add_separator()
        file_menu.add_command(label='Создать резервную копию…', command=self.backup)
        file_menu.add_command(label='Восстановить из резервной копии…', command=self.restore_backup)
        file_menu.add_separator()
        file_menu.add_command(label='Выход', command=self.on_close)
        menu.add_cascade(label='Файл', menu=file_menu)

        face_menu = tk.Menu(menu, tearoff=0)
        face_menu.add_command(label='Индексировать лица', command=self.index_faces)
        face_menu.add_command(label='Найти по лицу…', command=self.search_by_face)
        menu.add_cascade(label='Лица', menu=face_menu)
        self.config(menu=menu)

        toolbar = ttk.Frame(self, padding=(10, 8))
        toolbar.pack(fill='x')
        ttk.Label(toolbar, text='Поиск:').pack(side='left')
        self.search_var = tk.StringVar()
        e = ttk.Entry(toolbar, textvariable=self.search_var, width=42)
        e.pack(side='left', padx=(6, 6))
        e.bind('<Return>', lambda _e: self.load_rows())
        ttk.Button(toolbar, text='Найти', command=self.load_rows).pack(side='left')
        ttk.Button(toolbar, text='Сбросить', command=self.reset_search).pack(side='left', padx=(6, 0))
        ttk.Button(toolbar, text='Импорт фото', command=self.bulk_import_photos).pack(side='left', padx=(18, 0))
        ttk.Button(toolbar, text='Поиск по лицу', command=self.search_by_face).pack(side='left', padx=6)
        ttk.Button(toolbar, text='Новая карточка', command=self.new_record).pack(side='right')

        paned = ttk.Panedwindow(self, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        left, right = ttk.Frame(paned, padding=8), ttk.Frame(paned, padding=8)
        paned.add(left, weight=3)
        paned.add(right, weight=2)

        columns = ('archive_no', 'date', 'author', 'description')
        self.tree = ttk.Treeview(left, columns=columns, show='tree headings', selectmode='browse', height=18)
        self.tree.heading('#0', text='Фото')
        self.tree.heading('archive_no', text='Архивный №')
        self.tree.heading('date', text='Дата')
        self.tree.heading('author', text='Автор')
        self.tree.heading('description', text='Описание')
        self.tree.column('#0', width=105, minwidth=105, stretch=False)
        self.tree.column('archive_no', width=110, anchor='w')
        self.tree.column('date', width=95, anchor='w')
        self.tree.column('author', width=150, anchor='w')
        self.tree.column('description', width=420, anchor='w')
        yscroll = ttk.Scrollbar(left, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side='left', fill='both', expand=True)
        yscroll.pack(side='right', fill='y')
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        form = ttk.Frame(right)
        form.pack(fill='both', expand=True)
        self.entries = {}
        row = 0
        for key, label in [('archive_no', 'Архивный номер'), ('shot_date', 'Дата съёмки'),
                           ('location', 'Место съёмки'), ('author', 'Автор съёмки'),
                           ('source', 'Источник поступления')]:
            ttk.Label(form, text=label).grid(row=row, column=0, sticky='w', pady=(0, 3))
            ent = ttk.Entry(form)
            ent.grid(row=row + 1, column=0, sticky='ew', pady=(0, 8))
            self.entries[key] = ent
            row += 2

        ttk.Label(form, text='Описание фотографии').grid(row=row, column=0, sticky='w', pady=(0, 3)); row += 1
        self.description = tk.Text(form, height=7, wrap='word')
        self.description.grid(row=row, column=0, sticky='nsew', pady=(0, 8)); row += 1

        photo_frame = ttk.LabelFrame(form, text='Фотография', padding=6)
        photo_frame.grid(row=row, column=0, sticky='nsew', pady=(0, 8))
        self.preview = ttk.Label(photo_frame, text='Фото не выбрано', anchor='center')
        self.preview.pack(fill='both', expand=True)
        pb = ttk.Frame(photo_frame); pb.pack(fill='x', pady=(6, 0))
        ttk.Button(pb, text='Выбрать фото', command=self.choose_photo).pack(side='left')
        ttk.Button(pb, text='Открыть', command=self.open_photo).pack(side='left', padx=6)
        row += 1

        actions = ttk.Frame(form); actions.grid(row=row, column=0, sticky='ew')
        ttk.Button(actions, text='Сохранить', command=self.save_record).pack(side='left')
        ttk.Button(actions, text='Удалить', command=self.delete_record).pack(side='left', padx=6)
        ttk.Button(actions, text='Очистить', command=self.new_record).pack(side='left')
        form.columnconfigure(0, weight=1)
        form.rowconfigure(row - 1, weight=1)

        self.status_var = tk.StringVar(value='Готово')
        ttk.Label(self, textvariable=self.status_var, relief='sunken', anchor='w', padding=(6, 3)).pack(fill='x', side='bottom')

    def reset_search(self):
        self.search_var.set('')
        self.load_rows()

    def display_rows(self, rows, status=None):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.thumb_refs = {}
        for r in rows:
            desc = (r['description'] or '').replace('\n', ' ')
            if len(desc) > 80:
                desc = desc[:77] + '...'
            thumb_path = make_thumb(r['file_path'], r['id'])
            tkimg = None
            if thumb_path and ImageTk:
                try:
                    tkimg = ImageTk.PhotoImage(Image.open(thumb_path))
                    self.thumb_refs[str(r['id'])] = tkimg
                except Exception:
                    pass
            self.tree.insert('', 'end', iid=str(r['id']), text='', image=tkimg or '',
                             values=(r['archive_no'], r['shot_date'], r['author'], desc))
        self.status_var.set(status or f'Найдено записей: {len(rows)}')

    def load_rows(self):
        self.display_rows(self.db.search(self.search_var.get()))

    def on_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        row = self.db.get(int(sel[0]))
        if not row:
            return
        self.current_id = row['id']
        for key in self.entries:
            self.entries[key].delete(0, 'end'); self.entries[key].insert(0, row[key] or '')
        self.description.delete('1.0', 'end'); self.description.insert('1.0', row['description'] or '')
        self.current_image_path = row['file_path'] or ''
        self.show_preview(self.current_image_path)
        self.status_var.set(f"Открыта карточка ID {self.current_id}")

    def new_record(self):
        self.current_id = None; self.current_image_path = ''
        for ent in self.entries.values(): ent.delete(0, 'end')
        self.description.delete('1.0', 'end')
        self.preview.configure(image='', text='Фото не выбрано')
        self.preview_ref = None
        self.tree.selection_remove(self.tree.selection())
        self.status_var.set('Новая карточка')

    def choose_photo(self):
        path = filedialog.askopenfilename(title='Выберите фотографию',
            filetypes=[('Изображения', '*.jpg *.jpeg *.png *.bmp *.tif *.tiff'), ('Все файлы', '*.*')])
        if path:
            self.current_image_path = path; self.show_preview(path)

    def show_preview(self, path):
        self.preview_ref = None
        if not path or not os.path.exists(path):
            self.preview.configure(image='', text='Файл фотографии не найден'); return
        if Image is None:
            self.preview.configure(image='', text=os.path.basename(path)); return
        try:
            img = Image.open(path); img.thumbnail((440, 280))
            self.preview_ref = ImageTk.PhotoImage(img)
            self.preview.configure(image=self.preview_ref, text='')
        except Exception as exc:
            self.preview.configure(image='', text=f'Не удалось открыть фото\n{exc}')

    def open_photo(self):
        p = self.current_image_path
        if not p or not os.path.exists(p):
            messagebox.showwarning(APP_NAME, 'Фотография не выбрана или файл не найден.'); return
        try: os.startfile(p)
        except AttributeError:
            import subprocess; subprocess.Popen(['xdg-open', p])
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc))

    def collect_record(self):
        return {'archive_no': self.entries['archive_no'].get().strip(),
                'shot_date': self.entries['shot_date'].get().strip(),
                'location': self.entries['location'].get().strip(),
                'author': self.entries['author'].get().strip(),
                'source': self.entries['source'].get().strip(),
                'description': self.description.get('1.0', 'end').strip(),
                'file_path': self.current_image_path}

    def save_record(self):
        record = self.collect_record()
        if not record['archive_no']:
            messagebox.showwarning(APP_NAME, 'Укажите архивный номер.'); return
        if not record['file_path']:
            messagebox.showwarning(APP_NAME, 'Выберите фотографию.'); return
        old = self.db.get(self.current_id) if self.current_id else None
        try:
            p = record['file_path']
            if old and old['file_path'] == p and os.path.exists(p): archive_path = p
            elif str(Path(p).resolve()).startswith(str(PHOTOS_DIR.resolve())): archive_path = p
            else: archive_path = copy_into_archive(p)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f'Не удалось скопировать фотографию:\n{exc}'); return
        record['file_path'] = archive_path
        if self.current_id:
            self.db.update(self.current_id, record); msg = 'Карточка обновлена.'
        else:
            self.current_id = self.db.add(record); msg = 'Карточка добавлена.'
        self.current_image_path = archive_path
        self.load_rows()
        if self.tree.exists(str(self.current_id)):
            self.tree.selection_set(str(self.current_id)); self.tree.see(str(self.current_id))
        self.status_var.set(msg)

    def delete_record(self):
        if not self.current_id:
            messagebox.showinfo(APP_NAME, 'Сначала выберите карточку.'); return
        if not messagebox.askyesno(APP_NAME, 'Удалить выбранную карточку из базы?'): return
        row = self.db.delete(self.current_id)
        if row and row['file_path']:
            try:
                p = Path(row['file_path'])
                if p.exists() and str(p.resolve()).startswith(str(PHOTOS_DIR.resolve())): p.unlink()
            except Exception: pass
        self.new_record(); self.load_rows()

    def bulk_import_photos(self):
        paths = filedialog.askopenfilenames(title='Выберите фотографии для импорта',
            filetypes=[('Изображения', '*.jpg *.jpeg *.png *.bmp *.tif *.tiff'), ('Все файлы', '*.*')])
        if not paths: return
        prefix = simpledialog.askstring(APP_NAME, 'Префикс архивного номера (например Ф-):', initialvalue='Ф-')
        if prefix is None: return
        existing = self.db.all()
        next_num = len(existing) + 1
        ok, bad = 0, 0
        for src in paths:
            try:
                dest = copy_into_archive(src)
                rec = {'archive_no': f'{prefix}{next_num:06d}', 'description': Path(src).stem,
                       'shot_date': '', 'location': '', 'author': '', 'source': '', 'file_path': dest}
                self.db.add(rec); ok += 1; next_num += 1
            except Exception: bad += 1
            self.update_idletasks()
        self.load_rows()
        messagebox.showinfo(APP_NAME, f'Импорт завершён.\nДобавлено: {ok}\nОшибок: {bad}')

    def import_table(self):
        path = filedialog.askopenfilename(title='Excel или CSV', filetypes=[('Excel', '*.xlsx'), ('CSV', '*.csv'), ('Все файлы', '*.*')])
        if not path: return
        try:
            rows = []
            if path.lower().endswith('.xlsx'):
                if load_workbook is None: raise RuntimeError('Не установлен openpyxl.')
                ws = load_workbook(path, read_only=True, data_only=True).active
                data = list(ws.iter_rows(values_only=True))
                if not data: raise ValueError('Таблица пустая.')
                headers = [str(x or '').strip().lower() for x in data[0]]
                for vals in data[1:]: rows.append(dict(zip(headers, vals)))
            else:
                with open(path, 'r', encoding='utf-8-sig', newline='') as f:
                    rows = list(csv.DictReader(f))
            added, skipped = self._import_dict_rows(rows, Path(path).parent)
            self.load_rows(); messagebox.showinfo(APP_NAME, f'Импортировано: {added}\nПропущено: {skipped}')
        except Exception as exc: messagebox.showerror(APP_NAME, f'Ошибка импорта:\n{exc}')

    def _pick(self, row, aliases):
        lowered = {str(k).strip().lower(): v for k, v in row.items()}
        for a in aliases:
            if a in lowered and lowered[a] is not None: return str(lowered[a]).strip()
        return ''

    def _import_dict_rows(self, rows, base_dir):
        aliases = {
            'archive_no': ['архивный номер', 'архивный №', 'архивный номер.', 'archive_no', 'номер'],
            'description': ['описание', 'описание фотографии', 'description', 'содержание'],
            'shot_date': ['дата съемки', 'дата съёмки', 'дата', 'shot_date'],
            'location': ['место съемки', 'место съёмки', 'место', 'location'],
            'author': ['автор съемки', 'автор съёмки', 'автор', 'author'],
            'source': ['источник поступления', 'источник', 'source'],
            'file_path': ['фото', 'фотография', 'путь к фото', 'file_path', 'файл']}
        added = skipped = 0
        for row in rows:
            rec = {k: self._pick(row, v) for k, v in aliases.items()}
            if not rec['archive_no']: skipped += 1; continue
            fp = rec['file_path']
            if fp:
                p = Path(fp)
                if not p.is_absolute(): p = base_dir / p
                if p.exists(): rec['file_path'] = copy_into_archive(str(p))
                else: rec['file_path'] = ''
            self.db.add(rec); added += 1
        return added, skipped

    def import_access(self):
        path = filedialog.askopenfilename(title='База Microsoft Access', filetypes=[('Access', '*.accdb *.mdb')])
        if not path: return
        try:
            import pyodbc
            drivers = [d for d in pyodbc.drivers() if 'Access Driver' in d]
            if not drivers:
                raise RuntimeError('На компьютере не установлен Microsoft Access Database Engine / ODBC Driver.')
            conn = pyodbc.connect(r'DRIVER={' + drivers[-1] + r'};DBQ=' + path + ';')
            tables = [r.table_name for r in conn.cursor().tables(tableType='TABLE')]
            if not tables: raise RuntimeError('В базе не найдено таблиц.')
            table = tables[0]
            cur = conn.cursor(); cur.execute(f'SELECT * FROM [{table}]')
            cols = [d[0] for d in cur.description]
            rows = [dict(zip(cols, row)) for row in cur.fetchall()]
            conn.close()
            added, skipped = self._import_dict_rows(rows, Path(path).parent)
            self.load_rows(); messagebox.showinfo(APP_NAME, f'Таблица: {table}\nИмпортировано: {added}\nПропущено: {skipped}')
        except Exception as exc: messagebox.showerror(APP_NAME, f'Ошибка импорта Access:\n{exc}')

    def backup(self):
        target = filedialog.asksaveasfilename(title='Сохранить резервную копию', defaultextension='.zip',
            initialfile=f"PhotoArchive_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.zip", filetypes=[('ZIP', '*.zip')])
        if not target: return
        try:
            self.db.conn.commit()
            with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z:
                for p in DATA_DIR.rglob('*'):
                    if p.is_file(): z.write(str(p), str(p.relative_to(DATA_DIR)))
            messagebox.showinfo(APP_NAME, 'Резервная копия создана.')
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc))

    def restore_backup(self):
        source = filedialog.askopenfilename(title='Выберите резервную копию', filetypes=[('ZIP', '*.zip')])
        if not source: return
        if not messagebox.askyesno(APP_NAME, 'Текущая база будет заменена данными из резервной копии. Продолжить?'): return
        try:
            self.db.close()
            with zipfile.ZipFile(source, 'r') as z: z.extractall(DATA_DIR)
            self.db = CatalogDB(DB_PATH)
            self.new_record(); self.load_rows(); messagebox.showinfo(APP_NAME, 'Восстановление завершено.')
        except Exception as exc:
            self.db = CatalogDB(DB_PATH)
            messagebox.showerror(APP_NAME, str(exc))

    def index_faces(self):
        if not self.face.available:
            messagebox.showerror(APP_NAME, 'Поиск по лицу недоступен:\n' + self.face.error); return
        rows = self.db.all(); indexed = noface = errors = 0
        for i, r in enumerate(rows, 1):
            try:
                if r['file_path'] and os.path.exists(r['file_path']):
                    emb = self.face.embedding(r['file_path']); self.db.set_embedding(r['id'], emb); indexed += 1
                else: errors += 1
            except ValueError: noface += 1
            except Exception: errors += 1
            self.status_var.set(f'Индексация лиц: {i}/{len(rows)}'); self.update_idletasks()
        messagebox.showinfo(APP_NAME, f'Индексация завершена.\nЛиц: {indexed}\nБез лица: {noface}\nОшибок: {errors}')
        self.status_var.set('Готово')

    def search_by_face(self):
        if not self.face.available:
            messagebox.showerror(APP_NAME, 'Поиск по лицу недоступен:\n' + self.face.error); return
        path = filedialog.askopenfilename(title='Выберите фото лица для поиска',
            filetypes=[('Изображения', '*.jpg *.jpeg *.png *.bmp *.tif *.tiff')])
        if not path: return
        try: query = self.face.embedding(path)
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc)); return
        scored = []
        for r in self.db.embeddings():
            try:
                score = self.face.cosine(query, json.loads(r['embedding']))
                scored.append((score, r['id']))
            except Exception: pass
        scored.sort(reverse=True)
        ids = [pid for score, pid in scored if score >= 0.35][:100]
        rows = [self.db.get(pid) for pid in ids]
        self.display_rows(rows, f'Поиск по лицу: найдено {len(rows)} похожих фотографий')
        if not rows:
            messagebox.showinfo(APP_NAME, 'Совпадений не найдено. Сначала выполните «Лица → Индексировать лица».')

    def on_close(self):
        try: self.db.close()
        except Exception: pass
        self.destroy()


if __name__ == '__main__':
    PhotoCatalogApp().mainloop()

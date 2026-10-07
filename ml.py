
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_absolute_percentage_error
import warnings, os, threading, logging, re, pickle, copy, tempfile
import sqlite3, datetime, json, webbrowser, math
from tkinter import filedialog, messagebox, StringVar, IntVar, BooleanVar, DoubleVar
import tkinter as tk
import customtkinter as ctk
from scipy import stats

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GSHEETS_AVAILABLE = True
except ImportError:
    GSHEETS_AVAILABLE = False

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, filename='finscope.log', filemode='w',
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

CONFIG = {"min_data_points": 30, "max_lag": 14, "demo_data_periods": 180,
          "trial_days": 14, "app_name": "FinScope Pro", "version": "3.9.0",
          "company": "AI Business Solutions", "support_email": "support@aibusiness.ru",
          "r2_threshold": 0.1, "abc_a": 80.0, "abc_b": 95.0,
          "xyz_x": 10.0, "xyz_y": 25.0}

ABC_XYZ_PLAYBOOK = {
    'AX': "Непрерывная доступность. Страховой запас 3 дня. Автозаказ.",
    'AY': "Страховой запас 7 дней. Еженедельный контроль остатков.",
    'AZ': "Страховой запас 14 дней. Ручное управление закупками.",
    'BX': "Стандартный автозаказ. Страховой запас 5 дней.",
    'BY': "Еженедельный ручной пересмотр заказа.",
    'BZ': "Двухнедельный обзор; возможно объединение с позициями класса A.",
    'CX': "Минимальный заказ раз в месяц; рассмотреть drop-shipping.",
    'CY': "Рассмотреть вывод из ассортимента или объединение в пул.",
    'CZ': "Выводить из ассортимента или продавать под заказ."}

def sanitize(obj):
    if isinstance(obj, dict): return {k: sanitize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)): return [sanitize(v) for v in obj]
    if isinstance(obj, np.generic): return obj.item()
    return obj

class LicenseManager:
    @staticmethod
    def check_license():
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\FinScope")
            d = (datetime.datetime.now() -
                 datetime.datetime.fromisoformat(winreg.QueryValueEx(key, "FirstStart")[0])).days
            if d < CONFIG["trial_days"]:
                return True, f"Триал: осталось {CONFIG['trial_days'] - d} дн."
            return False, "Триал закончился. Приобретите лицензию."
        except Exception:
            try:
                import winreg
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\FinScope")
                winreg.SetValueEx(key, "FirstStart", 0, winreg.REG_SZ, datetime.datetime.now().isoformat())
                return True, f"Триал {CONFIG['trial_days']} дн. активирован"
            except Exception:
                if os.path.exists("trial_start.dat"):
                    with open("trial_start.dat") as f:
                        d0 = datetime.datetime.fromisoformat(f.read().strip())
                    d = (datetime.datetime.now() - d0).days
                    if d < CONFIG["trial_days"]:
                        return True, f"Триал: осталось {CONFIG['trial_days'] - d} дн."
                    return False, "Триал закончился."
                with open("trial_start.dat", "w") as f:
                    f.write(datetime.datetime.now().isoformat())
                return True, f"Триал {CONFIG['trial_days']} дн. активирован"

class HistoryDB:
    _lock = threading.Lock()
    EXPECTED_COLS = [('name', 'TEXT'), ('date', 'TEXT'), ('data', 'BLOB'),
                     ('notes', 'TEXT'), ('scenario', 'TEXT'), ('kpi_snapshot', 'TEXT')]
    def __init__(self, db_path="history.db"):
        self.db_path = db_path; self._local = threading.local()
        self._ensure_schema(self._get_conn())
    def _get_conn(self):
        if not hasattr(self._local, 'conn'):
            self._local.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        return self._local.conn
    def _ensure_schema(self, conn):
        conn.execute('''CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, date TEXT,
            data BLOB, notes TEXT, scenario TEXT, kpi_snapshot TEXT)''')
        existing = {r[1] for r in conn.execute("PRAGMA table_info(projects)")}
        for col, ctype in self.EXPECTED_COLS:
            if col not in existing:
                conn.execute(f"ALTER TABLE projects ADD COLUMN {col} {ctype}")
        conn.commit()
    def close(self):
        conn = getattr(self._local, 'conn', None)
        if conn is not None:
            try: conn.close()
            except Exception: pass
            self._local.conn = None
    def save_project(self, name, data, notes="", scenario="", kpi_snapshot=None):
        with self._lock:
            conn = self._get_conn()
            snap = json.dumps(sanitize(kpi_snapshot or {}), ensure_ascii=False)
            sql = "INSERT INTO projects (name,date,data,notes,scenario,kpi_snapshot) VALUES (?,?,?,?,?,?)"
            args = (name, datetime.datetime.now().isoformat(), pickle.dumps(data), notes, scenario, snap)
            try: cur = conn.execute(sql, args)
            except sqlite3.OperationalError:
                self._ensure_schema(conn); cur = conn.execute(sql, args)
            conn.commit(); return cur.lastrowid
    def get_projects(self):
        conn = self._get_conn(); self._ensure_schema(conn)
        return conn.execute("SELECT id,name,date,notes,scenario,kpi_snapshot FROM projects "
                            "ORDER BY date DESC").fetchall()
    def load_project(self, id):
        row = self._get_conn().execute("SELECT data FROM projects WHERE id=?", (id,)).fetchone()
        return pickle.loads(row[0]) if row else None
    def delete_project(self, id):
        with self._lock:
            conn = self._get_conn(); conn.execute("DELETE FROM projects WHERE id=?", (id,)); conn.commit()

def parse_dates(df, date_col):
    for fmt in [None, '%d.%m.%Y', '%Y-%m-%d', '%d-%m-%Y']:
        try:
            if fmt is None:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce', dayfirst=True)
            else:
                df[date_col] = pd.to_datetime(df[date_col], format=fmt, errors='coerce')
            if df[date_col].notna().any(): return df
        except Exception: pass
    raise ValueError("Не удалось распознать даты в файле")

def clean_currency_values(df, cols):
    for col in cols:
        if col in df.columns and df[col].dtype == object:
            s = df[col].astype(str).str.replace(r'[₽$€¥\s]', '', regex=True)
            mask_eu = s.str.contains(',', regex=False)
            s = s.where(~mask_eu, s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False))
            df[col] = pd.to_numeric(s, errors='coerce')
    return df

def parse_components_safe(s):
    if pd.isna(s) or not isinstance(s, str) or not s.strip(): return {}
    m = re.findall(r'([А-Яа-яA-Za-z0-9_]+)\s+([\d\.]+)', s)
    return {n: float(q) for n, q in m} if m else {}

def detect_columns(df):
    df.columns = [str(c).lstrip('\ufeff').strip() for c in df.columns]
    col_map, used = {}, set()
    def pick(t, keys, ex=()):
        for col in df.columns:
            if col in used: continue
            cl = str(col).lower()
            if any(k in cl for k in keys) and not any(e in cl for e in ex):
                col_map[t] = col; used.add(col); return
    pick('Дата', ['дата', 'date'])
    pick('Товар', ['товар', 'product', 'наименование', 'позиция'], ex=['сырь', 'компонент', 'состав'])
    pick('Компоненты', ['компонент', 'component', 'состав', 'рецептур'])
    pick('Цена за штуку', ['цена', 'price', 'стоимость', 'cost'])
    pick('Количество единиц', ['количество', 'quantity', 'кол-во', 'колво', 'шт', 'объем', 'объём', 'units'],
         ex=['цена', 'price', 'стоимость', 'cost'])
    if 'Дата' not in col_map:
        for col in df.columns:
            if col in used: continue
            try:
                pd.to_datetime(df[col], errors='raise'); col_map['Дата'] = col; used.add(col); break
            except Exception: pass
    if 'Количество единиц' not in col_map:
        for col in df.columns:
            if col in used: continue
            if pd.api.types.is_numeric_dtype(df[col]) and df[col].max() > 0:
                col_map['Количество единиц'] = col; used.add(col); break
    if 'Цена за штуку' not in col_map:
        for col in df.columns:
            if col in used: continue
            if pd.api.types.is_numeric_dtype(df[col]) and col != col_map.get('Количество единиц'):
                col_map['Цена за штуку'] = col; used.add(col); break
    return col_map

def load_production_data(fp):
    if isinstance(fp, pd.DataFrame): return fp.copy()
    ext = os.path.splitext(fp)[1].lower()
    if ext == '.csv':
        for enc in ('utf-8-sig', 'utf-8', 'cp1251'):
            try:
                with open(fp, 'r', encoding=enc) as f: head = f.readline()
                sep = ';' if head.count(';') > head.count(',') else ','
                df = pd.read_csv(fp, sep=sep, encoding=enc)
                if len(df.columns) > 1: return df
            except Exception: continue
        raise ValueError("Не удалось прочитать CSV")
    if ext == '.json':
        with open(fp, 'r', encoding='utf-8-sig') as f: payload = json.load(f)
        if isinstance(payload, dict) and 'production' in payload:
            return pd.DataFrame(payload['production'])
        if isinstance(payload, list): return pd.DataFrame(payload)
        return pd.json_normalize(payload)
    xl = pd.ExcelFile(fp)
    if 'Производство' in xl.sheet_names:
        return pd.read_excel(fp, sheet_name='Производство')
    best = max(xl.sheet_names, key=lambda s: pd.read_excel(fp, sheet_name=s).shape[1])
    return pd.read_excel(fp, sheet_name=best)

def load_raw_prices(fp):
    rp = {}
    if isinstance(fp, pd.DataFrame) or os.path.splitext(fp)[1].lower() not in ('.xlsx', '.xls'):
        return rp
    try:
        xl = pd.ExcelFile(fp)
        if 'Сырьё' not in xl.sheet_names: return rp
        df = pd.read_excel(fp, sheet_name='Сырьё')
        df = df.rename(columns={o: t for t, o in detect_columns(df).items()})
        if 'Цена за штуку' in df.columns:
            df = clean_currency_values(df, ['Цена за штуку'])
            nc = 'Товар' if 'Товар' in df.columns else df.columns[0]
            rp = {str(k): float(v) for k, v in df.set_index(nc)['Цена за штуку'].to_dict().items()}
    except Exception as e:
        logger.warning(f"Лист 'Сырьё': {e}")
    return rp

def load_user_raw_prices():
    if os.path.exists('raw_prices_user.json'):
        try:
            with open('raw_prices_user.json', encoding='utf-8') as f:
                return {str(k): float(v) for k, v in json.load(f).items()}
        except Exception:
            return {}
    return {}

class FinancialEngine:
    def __init__(self, p):
        self.fixed_costs_monthly = float(p.get('fixed_costs', 500000))
        self.opex_ratio = float(p.get('opex_ratio', 0.15))
        self.vat_rate = float(p.get('vat_rate', 0.20))
        self.profit_tax_rate = float(p.get('profit_tax_rate', 0.20))
        self.dso_days = float(p.get('dso_days', 15))
        self.dio_days = float(p.get('dio_days', 30))
        self.dpo_days = float(p.get('dpo_days', 20))
        self.price_cap = float(p.get('price_cap_pct', 15)) / 100
        self.marketing_share = float(p.get('marketing_pct', 2)) / 100
        self.price_change_cost = float(p.get('price_change_pct', 0.5)) / 100
        self.stress_pct = float(p.get('stress_pct', 20)) / 100
        self.opt_f = 1 + float(p.get('opt_pct', 15)) / 100
        self.pes_f = 1 - float(p.get('pes_pct', 15)) / 100
        self.elast_default = float(p.get('elast_default', 1.0))
        self.min_markup = float(p.get('min_markup_pct', 5)) / 100
        self.wc_threshold = float(p.get('wc_threshold', 2.0))
        self.ccc_threshold = float(p.get('ccc_threshold', 45))
        self.horizon = int(p.get('horizon_days', 7))
    def ccc(self): return self.dio_days + self.dso_days - self.dpo_days
    def unit_economics(self, avg_price_vat, cogs_per_unit):
        cogs = float(cogs_per_unit)
        pn = float(avg_price_vat) / (1 + self.vat_rate)
        vo = pn * self.opex_ratio
        vc = cogs + vo
        cm = pn - vc
        return {'cogs_per_unit': cogs, 'variable_cost_per_unit': vc, 'contribution_margin': cm,
                'cm_ratio': (cm / pn) if pn > 0 else 0, 'price_net_vat': pn, 'price_with_vat': float(avg_price_vat)}
    def break_even_monthly(self, fixed_m, cm, actual_m):
        if cm <= 0:
            return {'be_units_monthly': math.inf, 'is_profitable': False, 'margin_of_safety_pct': -100.0}
        be = fixed_m / cm
        mos = ((actual_m - be) / actual_m * 100) if actual_m > 0 else -100.0
        return {'be_units_monthly': float(be), 'is_profitable': bool(actual_m > be), 'margin_of_safety_pct': float(mos)}
    def net_profit(self, rev, cogs, vo, fixed, tax=True):
        g = rev - cogs; e = g - vo - fixed
        t = max(0.0, e * self.profit_tax_rate) if tax else 0.0
        n = e - t
        return {'revenue': float(rev), 'cogs': float(cogs), 'variable_opex': float(vo),
                'fixed_costs': float(fixed), 'gross_profit': float(g),
                'gross_margin_pct': float((g / rev * 100) if rev > 0 else 0),
                'ebitda': float(e), 'ebitda_margin_pct': float((e / rev * 100) if rev > 0 else 0),
                'tax': float(t), 'net_profit': float(n), 'net_margin_pct': float((n / rev * 100) if rev > 0 else 0)}
    def working_capital(self, mr, mc):
        i = mc / 30 * self.dio_days; r = mr / 30 * self.dso_days; p = mc / 30 * self.dpo_days
        return {'working_capital': float(i + r - p), 'avg_inventory': float(i),
                'avg_receivables': float(r), 'avg_payables': float(p), 'cash_conversion_cycle': float(self.ccc())}
    def cash_gap_warnings(self, mr, mc, growth=1.0):
        w = []
        ncf = mr - mc - mr * self.opex_ratio - self.fixed_costs_monthly
        d = self.working_capital(mr * growth, mc * growth)['working_capital'] - self.working_capital(mr, mc)['working_capital']
        if d > 0 and ncf > 0 and d / ncf > self.wc_threshold:
            w.append({'type': 'HIGH', 'msg': f"Рост потребует {d:,.0f} ₽ оборотного капитала "
                                             f"({d / ncf:.1f} мес. денежного потока; порог {self.wc_threshold:.1f})."})
        elif d > 0 and ncf <= 0:
            w.append({'type': 'CRITICAL', 'msg': f"Отрицательный денежный поток ({ncf:,.0f} ₽/мес) при росте "
                                                 f"оборотного капитала на {d:,.0f} ₽ — риск кассового разрыва."})
        if self.ccc() > self.ccc_threshold:
            w.append({'type': 'MEDIUM', 'msg': f"Денежный цикл {self.ccc():.0f} дн. выше порога "
                                               f"{self.ccc_threshold:.0f} дн."})
        return w

def optimal_price_step(ue, elast, cap):
    cur, c = ue['price_net_vat'], ue['variable_cost_per_unit']
    p_opt = c * elast / (elast - 1.0) if elast > 1.05 else cur * 1.10
    p_new = min(max(cur, p_opt), cur * (1 + cap))
    if p_new <= cur * 1.01:
        return cur, "Цена близка к оптимуму Лернера; повышение нецелесообразно"
    return p_new, f"Оптимум Лернера при эластичности {elast:.2f}, шаг ограничен {cap:.0%}"

def capital_required(a_type, old_u, new_u, ue, eng, rev_period, pm):
    dm = (new_u - old_u) / pm
    if dm > 0:
        return float(dm * ue['variable_cost_per_unit'] * max(0.0, eng.ccc()) / 30.0 +
                     eng.marketing_share * dm * ue['price_net_vat'])
    if 'цен' in a_type.lower():
        return float((rev_period / pm) * eng.price_change_cost)
    return 0.0

class DemandForecaster:
    def _build_features(self, df):
        df = df.copy()
        df['dayofweek'] = df['Дата'].dt.dayofweek
        df['month'] = df['Дата'].dt.month
        df['is_weekend'] = (df['dayofweek'] >= 5).astype(int)
        df['dow_sin'] = np.sin(2 * np.pi * df['dayofweek'] / 7)
        df['dow_cos'] = np.cos(2 * np.pi * df['dayofweek'] / 7)
        df['month_sin'] = np.sin(2 * np.pi * df['month'] / 12)
        df['month_cos'] = np.cos(2 * np.pi * df['month'] / 12)
        past = df['Количество'].shift(1)
        df['ma7'] = past.rolling(7, min_periods=1).mean()
        df['ma14'] = past.rolling(14, min_periods=1).mean()
        df['ma30'] = past.rolling(30, min_periods=1).mean()
        df['std7'] = past.rolling(7, min_periods=1).std().fillna(0)
        df['trend'] = past.rolling(14, min_periods=1).mean() - past.rolling(30, min_periods=1).mean()
        return df.bfill().ffill()

    def fit_predict(self, df_prod, eng, scenario, progress_callback=None, stop_flag=None):
        res = {}; prods = df_prod['Товар'].unique(); tot = len(prods)
        for idx, prod in enumerate(prods):
            if stop_flag and stop_flag(): raise InterruptedError("Расчёт прерван")
            if progress_callback:
                progress_callback(f"Прогноз спроса: {prod}...", 30 + int(50 * idx / tot))
            daily = df_prod[df_prod['Товар'] == prod].groupby('Дата')['Количество единиц'] \
                .sum().reset_index().rename(columns={'Количество единиц': 'Количество'})
            daily = daily.set_index('Дата').asfreq('D').fillna(0).reset_index()
            hist_std = float(daily['Количество'].std())
            def fb():
                # v3.9.0: честный fallback — среднее и ФАКТИЧЕСКОЕ sigma истории,
                # точность не выдумывается (None)
                return {'historical': daily,
                        'forecast_dates': pd.date_range(start=daily['Дата'].iloc[-1] + pd.Timedelta(days=1),
                                                        periods=eng.horizon),
                        'forecast_units': [float(daily['Количество'].mean())] * eng.horizon,
                        'forecast_std': hist_std, 'accuracy': None, 'method': 'среднее историческое',
                        'feature_importance': {}}
            if len(daily) < CONFIG['min_data_points']:
                res[prod] = fb(); continue
            dfF = self._build_features(daily)
            ml = min(CONFIG['max_lag'], len(dfF) - eng.horizon - 1)
            for l in range(1, ml + 1): dfF[f'lag_{l}'] = dfF['Количество'].shift(l)
            dfF = dfF.dropna().reset_index(drop=True)
            if ml < 1 or len(dfF) < eng.horizon + ml:
                res[prod] = fb(); continue
            fcols = [f'lag_{l}' for l in range(1, ml + 1)] + \
                    ['dow_sin', 'dow_cos', 'month_sin', 'month_cos', 'is_weekend', 'ma7', 'ma14', 'ma30', 'std7', 'trend']
            X, y = dfF[fcols].values, dfF['Количество'].values
            sp = -eng.horizon if len(dfF) > eng.horizon + 10 else int(len(dfF) * 0.8)
            Xtr, Xte, ytr, yte = X[:sp], X[sp:], y[:sp], y[sp:]
            sc = StandardScaler(); Xtr_s, Xte_s = sc.fit_transform(Xtr), sc.transform(Xte)
            models = [('LR', LinearRegression(), True),
                      ('RF', RandomForestRegressor(n_estimators=100, max_depth=8, random_state=42, n_jobs=-1), False),
                      ('GB', GradientBoostingRegressor(n_estimators=100, learning_rate=0.1, max_depth=4, random_state=42), False)]
            preds, trained = [], []
            for nm, md, ns in models:
                try:
                    if ns: md.fit(Xtr_s, ytr); p = np.maximum(md.predict(Xte_s), 0)
                    else: md.fit(Xtr, ytr); p = np.maximum(md.predict(Xte), 0)
                    preds.append(p); trained.append((nm, md, ns))
                except Exception as e: logger.warning(f"{nm}: {e}")
            if not trained:
                res[prod] = fb(); continue
            if len(yte) > 0:
                mapes = []
                for p in preds:
                    try: mapes.append(mean_absolute_percentage_error(yte, p) * 100)
                    except Exception: mapes.append(100)
                inv = [1.0 / (m + 1e-6) for m in mapes]
                wts = [w / sum(inv) for w in inv]
                ens = sum(p * w for p, w in zip(preds, wts))
                rstd = float(np.std(yte - ens))
                try: acc = float(max(0.0, 100 - mean_absolute_percentage_error(yte, ens) * 100))
                except Exception: acc = None
            else:
                wts = [1.0 / len(trained)] * len(trained); rstd = hist_std; acc = None
            lags = daily['Количество'].values[-ml:]
            fdates = pd.date_range(start=daily['Дата'].iloc[-1] + pd.Timedelta(days=1), periods=eng.horizon)
            fpreds = []
            for st in range(eng.horizon):
                d = fdates[st]
                ft = {'dow_sin': np.sin(2*np.pi*d.dayofweek/7), 'dow_cos': np.cos(2*np.pi*d.dayofweek/7),
                      'month_sin': np.sin(2*np.pi*d.month/12), 'month_cos': np.cos(2*np.pi*d.month/12),
                      'is_weekend': 1 if d.dayofweek >= 5 else 0, 'ma7': np.mean(lags[-7:]),
                      'ma14': np.mean(lags[-14:]) if len(lags) >= 14 else np.mean(lags),
                      'ma30': np.mean(lags), 'std7': np.std(lags[-7:]),
                      'trend': (np.mean(lags[-14:]) if len(lags) >= 14 else np.mean(lags)) - np.mean(lags)}
                for l in range(1, ml + 1): ft[f'lag_{l}'] = lags[-l]
                fv = np.array([ft[c] for c in fcols]).reshape(1, -1)
                pr = sum(w * md.predict(sc.transform(fv) if ns else fv)[0] for (nm, md, ns), w in zip(trained, wts))
                pr = max(0.0, float(pr)); fpreds.append(pr); lags = np.append(lags[1:], pr)
            fac = {'optimistic': eng.opt_f, 'realistic': 1.0, 'pessimistic': eng.pes_f}.get(scenario, 1.0)
            fi = {c: 0.0 for c in fcols}
            for (nm, md, _), w in zip(trained, wts):
                if hasattr(md, 'feature_importances_'):
                    for c, v in zip(fcols, md.feature_importances_): fi[c] += float(v) * w
            res[prod] = {'historical': daily, 'forecast_dates': fdates,
                         'forecast_units': [float(p) * fac for p in fpreds], 'forecast_std': rstd,
                         'accuracy': round(acc, 1) if acc is not None else None,
                         'method': 'ансамбль LR+RF+GBR', 'feature_importance': fi}
        return res

class PortfolioAnalyzer:
    """Строгая классификация портфеля: доля в выручке x прогнозный рост."""
    @staticmethod
    def classify(pd_, fc, mg=None):
        gr = {}
        for prod, data in pd_.items():
            f = fc.get(prod, {}); h = f.get('historical'); g = 0.0
            if h is not None and len(h) >= 30:
                rec = float(h['Количество'].tail(30).mean())
                if rec > 0: g = (float(np.mean(f.get('forecast_units', [0]))) - rec) / rec
            gr[prod] = g
        if mg is None: mg = max(0.0, float(np.median(list(gr.values()))))
        out = {}; tr = sum(p['pl']['revenue'] for p in pd_.values()) or 1
        for prod, data in pd_.items():
            sh = data['pl']['revenue'] / tr; g = gr[prod]
            hs, hg = sh > 0.15, g > mg
            if hs and hg: q, s = "Лидер (рост)", "Инвестировать в масштабирование"
            elif hs: q, s = "Лидер (стабильность)", "Удерживать маржу, извлекать прибыль"
            elif hg: q, s = "Перспективный", "Оценить целесообразность инвестиций в долю"
            else: q, s = "Низкий приоритет", "Минимизировать вложения"
            out[prod] = {'share_pct': round(float(sh)*100, 1), 'growth_pct': round(float(g)*100, 1),
                         'quadrant': q, 'strategy': s, 'threshold_pct': round(mg*100, 1)}
        return out

def abc_xyz_analysis(df_prod, prod_agg):
    if prod_agg.empty: return prod_agg
    df = prod_agg.copy().sort_values('total_revenue', ascending=False)
    df['revenue_cum_pct'] = df['total_revenue'].cumsum() / (df['total_revenue'].sum() or 1) * 100
    df['abc_class'] = 'C'
    df.loc[df['revenue_cum_pct'] <= CONFIG['abc_a'], 'abc_class'] = 'A'
    df.loc[(df['revenue_cum_pct'] > CONFIG['abc_a']) & (df['revenue_cum_pct'] <= CONFIG['abc_b']), 'abc_class'] = 'B'
    xyz = {}
    for p in df['Товар'].unique():
        d = df_prod[df_prod['Товар'] == p].groupby('Дата')['Количество единиц'].sum()
        if len(d) > 5 and d.mean() > 0:
            cov = float(d.std()) / float(d.mean()) * 100
            xyz[p] = 'X' if cov < CONFIG['xyz_x'] else 'Y' if cov < CONFIG['xyz_y'] else 'Z'
        else: xyz[p] = 'Z'
    df['xyz_class'] = df['Товар'].map(xyz); df['abc_xyz'] = df['abc_class'] + df['xyz_class']
    return df

def calculate_elasticity(df_prod):
    out = {}
    for p in df_prod['Товар'].unique():
        d = df_prod[df_prod['Товар'] == p].groupby('Дата').agg(
            q=('Количество единиц', 'sum'), p=('Цена за штуку', 'mean')).reset_index()
        d = d[(d.q > 0) & (d.p > 0)]
        if len(d) < 10:
            out[p] = (None, 'manual'); continue
        try:
            r = stats.linregress(np.log(d.p), np.log(d.q))
            out[p] = (round(float(abs(r.slope)), 2), 'data') if r.rvalue**2 > CONFIG['r2_threshold'] else (None, 'manual')
        except Exception:
            out[p] = (None, 'manual')
    return out

def sensitivity_analysis(pe, actions, pm, stress):
    sc = {'base': 1.0, f'спрос −{stress*100:.0f}%': 1 - stress, f'спрос +{stress*100:.0f}%': 1 + stress}
    res = {}
    for nm, k in sc.items():
        t = 0.0
        for a in actions:
            e = pe[a['Товар']]
            u = a['Новые_шт'] * k
            rv = u * a['Новая_цена']
            t += (rv - u * e['unit_economics']['cogs_per_unit'] - rv * e['opex_ratio'] -
                  e['fixed_share']) * (1 - e['tax_rate'])
        res[nm] = round(t / pm, 2)
    v = list(res.values()); mn, mx = min(v), max(v)
    sp = (mx - mn) / abs(mn) * 100 if mn != 0 else 0
    return {'scenarios': res, 'spread_pct': round(sp, 1),
            'stability': "Устойчивый" if sp < 20 else "Зависит от спроса" if sp < 50 else "Хрупкий"}

def scenario_comparison_table(rep):
    eng = FinancialEngine(rep.get('engine_params', {}))
    pe = rep['product_economics']
    rows = []
    for sname, f in [("Оптимистичный", eng.opt_f), ("Реалистичный", 1.0), ("Пессимистичный", eng.pes_f)]:
        no_act, with_plan = 0.0, 0.0
        for p, e in pe.items():
            a = next((x for x in rep['actions'] if x['Товар'] == p), None)
            pn = e['unit_economics']['price_net_vat']; cu = e['cogs_per_unit']
            cm = pn - cu - pn * eng.opex_ratio
            no_act += (e['units_period'] * f * cm - e['fixed_share']) * (1 - e['tax_rate'])
            pp = a['Новая_цена'] if a else pn
            uu = (a['Новые_шт'] if a else e['units_period']) * f
            cm_p = pp - cu - pp * eng.opex_ratio
            with_plan += (uu * cm_p - e['fixed_share']) * (1 - e['tax_rate'])
        rows.append({'Сценарий': sname, 'Без изменений': round(no_act),
                     'С планом': round(with_plan), 'Эффект': round(with_plan - no_act)})
    return rows

class ImportWizard(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent; self.title("Мастер импорта"); self.geometry("750x650")
        ctk.set_appearance_mode("light"); self.df = None; self.result = None
        ctk.CTkLabel(self, text="Мастер импорта данных", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        s = ctk.CTkFrame(self, fg_color="transparent"); s.pack(pady=10)
        ctk.CTkButton(s, text="Excel/CSV", command=self.load_file).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(s, text="Google Sheets", command=self.load_gsheets).pack(side=tk.LEFT, padx=5)
        self.url_entry = ctk.CTkEntry(self, placeholder_text="URL Google Sheets", width=400); self.url_entry.pack(pady=5)
        self.preview_text = ctk.CTkTextbox(self, height=150, font=ctk.CTkFont(size=10))
        self.preview_text.pack(fill=tk.BOTH, padx=10, pady=10, expand=True)
        self.mw = {}
        for col in ["Дата", "Товар", "Компоненты", "Количество единиц", "Цена за штуку"]:
            r = ctk.CTkFrame(self, fg_color="transparent"); r.pack(fill=tk.X, pady=2)
            ctk.CTkLabel(row if False else r, text=f"{col}:", width=160, anchor="w").pack(side=tk.LEFT, padx=5)
            c = ctk.CTkComboBox(r, values=["(не выбрано)"], width=220); c.pack(side=tk.LEFT, padx=5)
            self.mw[col] = c
        b = ctk.CTkFrame(self, fg_color="transparent"); b.pack(pady=10)
        ctk.CTkButton(b, text="Автоопределить", command=self.auto_detect).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Применить маппинг", command=self.apply_mapping).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Завершить", command=self.finish, fg_color="#0057b3").pack(side=tk.LEFT, padx=5)
    def auto_detect(self):
        if self.df is None: messagebox.showerror("Ошибка", "Сначала загрузите данные"); return
        cm = detect_columns(self.df)
        for c in self.mw.values(): c.configure(values=["(не выбрано)"] + list(self.df.columns))
        for k, c in self.mw.items(): c.set(cm.get(k, "(не выбрано)"))
    def load_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("JSON", "*.json")])
        if not p: return
        try:
            self.df = load_production_data(p)
            self.preview_text.delete("0.0", tk.END)
            self.preview_text.insert("0.0", self.df.head(10).to_string()); self.auto_detect()
        except Exception as e: messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
    def load_gsheets(self):
        if not GSHEETS_AVAILABLE: messagebox.showerror("Ошибка", "pip install gspread oauth2client"); return
        u = self.url_entry.get()
        if not u: messagebox.showerror("Ошибка", "Введите URL"); return
        try:
            if not os.path.exists("credentials.json"):
                messagebox.showerror("Ошибка", "credentials.json не найден."); return
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            cr = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
            data = gspread.authorize(cr).open_by_url(u).get_worksheet(0).get_all_values()
            if not data: raise ValueError("Таблица пуста")
            self.df = pd.DataFrame(data[1:], columns=data[0])
            self.preview_text.delete("0.0", tk.END)
            self.preview_text.insert("0.0", self.df.head(10).to_string()); self.auto_detect()
        except Exception as e: messagebox.showerror("Ошибка", f"Google Sheets: {e}")
    def apply_mapping(self):
        if self.df is None: messagebox.showerror("Ошибка", "Сначала загрузите данные"); return
        mp = {}
        for k, c in self.mw.items():
            v = c.get()
            if v == "(не выбрано)": messagebox.showerror("Ошибка", f"Выберите колонку для '{k}'"); return
            mp[k] = v
        try:
            r = self.df.rename(columns={v: k for k, v in mp.items()})
            r = parse_dates(r, "Дата"); r = clean_currency_values(r, ["Количество единиц", "Цена за штуку"])
            self.result = r; messagebox.showinfo("Успех", "Маппинг применён.")
        except Exception as e: messagebox.showerror("Ошибка", str(e))
    def finish(self):
        if self.result is not None:
            self.parent.import_result = self.result
            self.parent.file_path.set("Импортировано через мастер"); self.destroy()
        else: messagebox.showerror("Ошибка", "Сначала примените маппинг")

def generate_demo_data():
    np.random.seed(42)
    dates = pd.date_range(start='2025-01-01', periods=CONFIG["demo_data_periods"], freq='D')
    prods = [f"Товар_{i}" for i in range(1, 6)]
    comp = {"Товар_1": "Сырьё_A 2 Сырьё_B 3", "Товар_2": "Сырьё_A 1 Сырьё_C 4",
            "Товар_3": "Сырьё_B 2 Сырьё_C 2", "Товар_4": "Сырьё_A 3 Сырьё_D 1",
            "Товар_5": "Сырьё_C 2 Сырьё_D 3"}
    prices = {"Сырьё_A": 120, "Сырьё_B": 80, "Сырьё_C": 45, "Сырьё_D": 95}
    bp = {"Товар_1": 850, "Товар_2": 650, "Товар_3": 520, "Товар_4": 950, "Товар_5": 480}
    bd = {"Товар_1": 25, "Товар_2": 40, "Товар_3": 60, "Товар_4": 15, "Товар_5": 35}
    rows = []
    for d in dates:
        for p in prods:
            se = 1 + 0.25 * np.sin(2 * np.pi * d.dayofyear / 365)
            tr = 1 + 0.0005 * (d - dates[0]).days
            we = 1.2 if d.dayofweek >= 5 else 1.0
            u = max(1, int(bd[p] * se * tr * we * np.random.normal(1.0, 0.15)))
            rows.append([d, p, comp[p], u, bp[p] * np.random.uniform(0.95, 1.05)])
    dfp = pd.DataFrame(rows, columns=["Дата", "Товар", "Компоненты", "Количество единиц", "Цена за штуку"])
    dfr = pd.DataFrame(list(prices.items()), columns=["Товар", "Цена за штуку"])
    fd, path = tempfile.mkstemp(suffix='.xlsx'); os.close(fd)
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        dfr.to_excel(w, sheet_name="Сырьё", index=False)
        dfp.to_excel(w, sheet_name="Производство", index=False)
    return path

def generate_report_data(file_path, target_margin=20, increase_pct=20, scenario="realistic",
                         progress_callback=None, stop_flag=None, params=None):
    logger.info("Генерация отчёта v3.9.0")
    if progress_callback: progress_callback("Проверка данных...", 5)
    params = dict(params or {})
    eng = FinancialEngine(params)

    df0 = load_production_data(file_path)
    df = df0.rename(columns={o: t for t, o in detect_columns(df0).items()})
    miss = [c for c in ["Дата", "Товар", "Количество единиц", "Цена за штуку"] if c not in df.columns]
    if miss: raise ValueError(f"Отсутствуют колонки: {miss}. Используйте мастер импорта.")
    n_in = len(df)
    df = parse_dates(df, "Дата"); df = clean_currency_values(df, ["Количество единиц", "Цена за штуку"])
    df = df.dropna(subset=["Дата", "Количество единиц", "Цена за штуку"])
    df = df[(df["Количество единиц"] >= 0) & (df["Цена за штуку"] > 0)]
    n_dropped = n_in - len(df)
    if df.empty: raise ValueError("После очистки не осталось корректных строк.")
    if "Компоненты" not in df.columns: df["Компоненты"] = ""

    all_raw = set()
    for _, r in df.iterrows():
        all_raw.update(parse_components_safe(r.get('Компоненты', '')).keys())
    raw_prices = load_raw_prices(file_path)
    n_file = len(raw_prices)
    raw_prices.update(params.get('user_raw_prices', {}))
    n_user = len(raw_prices) - n_file
    missing_raw = sorted(all_raw - set(raw_prices))
    if all_raw and missing_raw:
        raise ValueError("Не указаны цены сырья: " + ", ".join(missing_raw) +
                         ".\nОткройте «Цены сырья» и заполните цены, либо добавьте лист «Сырьё» в файл.")
    if progress_callback: progress_callback("Данные загружены", 10)

    cl, rl = [], []
    for _, r in df.iterrows():
        if stop_flag and stop_flag(): raise InterruptedError("Расчёт прерван пользователем")
        cp = parse_components_safe(r.get('Компоненты', ''))
        cl.append(float(sum(q * raw_prices.get(c, 0) for c, q in cp.items())))
        rl.append(float(r['Цена за штуку']) * float(r['Количество единиц']) / (1 + eng.vat_rate))
    df['Себестоимость'], df['Выручка'] = cl, rl

    pa = df.groupby('Товар').agg(
        total_units=('Количество единиц', 'sum'), total_revenue=('Выручка', 'sum'),
        total_cogs=('Себестоимость', 'sum'), avg_price=('Цена за штуку', 'mean'),
        days_active=('Дата', 'nunique'),
        recipe=('Компоненты', lambda x: parse_components_safe(x.iloc[0]) if len(x) else {})).reset_index()
    pa['avg_daily_units'] = pa['total_units'] / pa['days_active'].clip(lower=1)
    pa['cogs_per_unit'] = pa['total_cogs'] / pa['total_units'].clip(lower=1)
    dr = max(1, (df['Дата'].max() - df['Дата'].min()).days + 1); pm = dr / 30.0

    if progress_callback: progress_callback("Эластичность, ABC-XYZ...", 15)
    elast_map = calculate_elasticity(df)
    pa['elasticity'] = pa['Товар'].map(lambda p: elast_map.get(p, (None, 'manual'))[0])
    pa['elast_src'] = pa['Товар'].map(lambda p: elast_map.get(p, (None, 'manual'))[1])
    pa = abc_xyz_analysis(df, pa)

    if progress_callback: progress_callback("Юнит-экономика и P&L...", 25)
    tra = float(pa['total_revenue'].sum()) or 1
    pe = {}
    for _, r in pa.iterrows():
        p = r['Товар']; sh = float(r['total_revenue']) / tra
        fm = eng.fixed_costs_monthly * sh; fp = fm * pm
        ue = eng.unit_economics(r['avg_price'], r['cogs_per_unit'])
        am = float(r['avg_daily_units']) * 30
        be = eng.break_even_monthly(fm, ue['contribution_margin'], am)
        pl = eng.net_profit(r['total_revenue'], r['total_cogs'], r['total_revenue'] * eng.opex_ratio, fp)
        pe[p] = {'unit_economics': ue, 'break_even': be, 'pl': pl, 'fixed_share': float(fp),
                 'opex_ratio': eng.opex_ratio, 'tax_rate': eng.profit_tax_rate,
                 'revenue_share': sh * 100, 'units_period': float(r['total_units']),
                 'revenue_period': float(r['total_revenue']), 'cogs_period': float(r['total_cogs']),
                 'cogs_per_unit': float(r['cogs_per_unit']), 'avg_price_vat': float(r['avg_price'])}

    if progress_callback: progress_callback("Прогноз спроса (ML)...", 30)
    fc = DemandForecaster().fit_predict(df, eng, scenario,
                                        progress_callback=progress_callback, stop_flag=stop_flag)
    if progress_callback: progress_callback("Классификация портфеля...", 70)
    bcg = PortfolioAnalyzer.classify(pe, fc)

    if progress_callback: progress_callback("Построение рекомендаций...", 80)
    inc = 1 + increase_pct / 100.0; tn = target_margin / 100.0
    actions, pf = [], []
    for _, r in pa.iterrows():
        p = r['Товар']; e = pe[p]; pl, ue, be = e['pl'], e['unit_economics'], e['break_even']
        el_raw, el_src = r['elasticity'], r['elast_src']
        el = float(el_raw) if el_raw is not None else eng.elast_default
        el_note = "эластичность из данных" if el_src == 'data' else f"эластичность {el:.2f} задана в настройках"
        q = bcg.get(p, {}).get('quadrant', '')
        ou, op = float(r['total_units']), ue['price_net_vat']
        cm = ue['contribution_margin']; am = float(r['avg_daily_units']) * 30; bm = be['be_units_monthly']

        if cm <= 0:
            at = "Вывести из ассортимента"; nu, np_ = ou * 0.1, op
            rs = f"Отрицательный маржинальный доход ({el_note})"
        elif q == "Лидер (рост)":
            at = "Увеличить выпуск"; nu, np_ = ou * inc, op
            rs = f"Рост спроса {bcg[p]['growth_pct']}% при доле {bcg[p]['share_pct']}%"
        elif q == "Лидер (стабильность)":
            if el_src == 'data' and el < 1.0:
                np_, why = optimal_price_step(ue, el, eng.price_cap); at = "Повысить цену"; nu = ou
                rs = f"Неэластичный спрос. {why} ({el_note})"
            else:
                at = "Поддерживать объём"; nu, np_ = ou, op
                rs = f"Стабильный лидер: удерживать маржу ({el_note})"
        elif q == "Низкий приоритет" and pl['net_margin_pct'] >= tn * 100:
            if el_src == 'data' and el < 1.0:
                np_, why = optimal_price_step(ue, el, eng.price_cap); at = "Повысить цену"; nu = ou
                rs = f"Прибыльная ниша, неэластичный спрос. {why} ({el_note})"
            else:
                at = "Поддерживать объём"; nu, np_ = ou, op
                rs = f"Прибыльная ниша: держать объём и маржу ({el_note})"
        elif q == "Низкий приоритет" and math.isfinite(bm) and am < bm:
            nu = min(bm * 1.05 * pm, ou); at = "Вывести на безубыточность"; np_ = op
            rs = f"Продажи {am:.0f} шт/мес ниже точки безубыточности {bm:.0f} шт/мес"
        elif q == "Перспективный" and pl['net_margin_pct'] > 0:
            at = "Умеренный рост"; nu, np_ = ou * (1 + increase_pct / 200.0), op
            rs = "Растущий рынок при положительной марже"
        elif el_src == 'data' and el < 1.0 and pl['net_margin_pct'] < tn * 100:
            np_, why = optimal_price_step(ue, el, eng.price_cap); at = "Повысить цену"; nu = ou
            rs = f"Маржа {pl['net_margin_pct']:.1f}% ниже целевой. {why} ({el_note})"
        elif pl['net_margin_pct'] < 0:
            at = "Сократить выпуск"; nu, np_ = ou * 0.3, op; rs = "Убыточная позиция"
        else:
            at = "Поддерживать объём"; nu, np_ = ou, op; rs = f"Показатели в норме ({el_note})"

        nu = max(0.0, float(nu)); np_ = max(float(np_), ue['cogs_per_unit'] * (1 + eng.min_markup))
        su = (nu / ou) if ou > 0 else 0.0
        prf = (np_ / op) if op > 0 else 1.0
        new_rev = float(r['total_revenue']) * su * prf
        new_cogs = float(r['total_cogs']) * su
        npl = eng.net_profit(new_rev, new_cogs, new_rev * eng.opex_ratio, e['fixed_share'])

        cap = capital_required(at, ou, nu, ue, eng, float(r['total_revenue']), pm)
        dm = (npl['net_profit'] - pl['net_profit']) / pm
        pb = cap / dm if dm > 0 else math.inf
        rk = "Низкий"
        if el > 1.5 and np_ > op * 1.01: rk = "Высокий (эластичный спрос)"
        elif 'Сократить' in at or 'Вывести' in at: rk = "Средний (потеря доли рынка)"

        rd = {}
        if nu > ou:
            dmn = (nu - ou) / pm
            rd = {c: int(round(q * dmn)) for c, q in (r['recipe'] or {}).items()}

        pf.append({'Товар': p, 'ABC-XYZ': r.get('abc_xyz', '??'), 'Класс': q,
                   'Текущая чистая прибыль': round(float(pl['net_profit']), 2),
                   'Прогнозная чистая прибыль': round(float(npl['net_profit']), 2),
                   'Изменение': round(float(npl['net_profit'] - pl['net_profit']), 2),
                   'Текущая маржа %': round(float(pl['net_margin_pct']), 1),
                   'Прогнозная маржа %': round(float(npl['net_margin_pct']), 1),
                   'Эластичность': round(el, 2), 'Источник эластичности': el_note,
                   'Действие': at, 'Обоснование': rs})
        actions.append({'Товар': p, 'Тип': at, 'Новые_шт': int(round(nu)),
                        'Новая_цена': round(float(np_), 2), 'Обоснование': rs, 'elasticity': el,
                        'raw_delta': rd,
                        'roi': {'capital_required': round(float(cap), 2),
                                'delta_profit_month': round(float(dm), 2),
                                'payback_months': round(float(pb), 1) if pb != math.inf else 'н/д',
                                'risk_level': rk}})

    tnow = float(sum(x['Текущая чистая прибыль'] for x in pf))
    tfut = float(sum(x['Прогнозная чистая прибыль'] for x in pf))
    chg = ((tfut - tnow) / abs(tnow)) * 100 if tnow != 0 else 0

    mrev = float(pa['total_revenue'].sum()) / pm
    mcogs = float(pa['total_cogs'].sum()) / pm
    grw = inc if any(a['Новые_шт'] * inc > a['Новые_шт'] + 1 for a in actions) else 1.0
    cw = eng.cash_gap_warnings(mrev, mcogs, grw)
    if not all_raw:
        cw.insert(0, {'type': 'MEDIUM', 'msg': "Рецептуры не найдены: себестоимость сырья не учитывается "
                                               "(маржа = цена минус переменные OPEX и постоянные расходы)."})

    daily = df.groupby('Дата')['Выручка'].sum().reset_index().sort_values('Дата')
    cmr = {}
    for _, r in df.iterrows():
        for c, q in parse_components_safe(r.get('Компоненты', '')).items():
            cmr[c] = cmr.get(c, 0) + q / pm
    ad = {}
    for a in actions:
        for c, d in a['raw_delta'].items(): ad[c] = ad.get(c, 0) + d
    rr = [{'Сырьё': c, 'Текущее (мес)': int(round(v)), 'Изменение': ad.get(c, 0),
           'Новое': int(round(v + ad.get(c, 0))), 'Цена': round(float(raw_prices.get(c, 0)), 2)}
          for c, v in sorted(cmr.items())]

    st = sensitivity_analysis(pe, actions, pm, eng.stress_pct)
    ar = [{'Товар': r['Товар'], 'Класс': r.get('abc_xyz', '??'),
           'Рекомендация по запасам': ABC_XYZ_PLAYBOOK.get(r.get('abc_xyz'), "Требуется анализ")}
          for _, r in pa.iterrows()]
    if progress_callback: progress_callback("Готово", 100)

    accs = [f['accuracy'] for f in fc.values() if f['accuracy'] is not None]
    aacc = float(np.mean(accs)) if accs else None
    svar = sum(f.get('forecast_std', 0) ** 2 for f in fc.values())
    ap = float(pa['total_revenue'].sum()) / float(pa['total_units'].sum()) if pa['total_units'].sum() > 0 else 0
    cih = 1.96 * math.sqrt(svar) * ap
    if fc:
        fst = next(iter(fc.values())); fdates = fst['forecast_dates']
        un = [0.0] * len(fdates)
        for f in fc.values():
            for i, u in enumerate(f['forecast_units']): un[i] += u
        frev = [u * ap for u in un]
    else:
        fdates = pd.date_range(start=df['Дата'].max() + pd.Timedelta(days=1), periods=eng.horizon)
        frev = [0.0] * eng.horizon
    fi = {}
    for f in fc.values():
        for k, v in f.get('feature_importance', {}).items(): fi[k] = fi.get(k, 0) + v
    if fi:
        mx = max(fi.values()) or 1; fi = {k: v / mx for k, v in fi.items()}

    thr = next(iter(bcg.values()), {}).get('threshold_pct', 0.0)
    assumptions = {
        'НДС': f"{eng.vat_rate*100:.0f}% (цены продажи считаются с НДС)",
        'Налог на прибыль': f"{eng.profit_tax_rate*100:.0f}%",
        'DSO / DIO / DPO': f"{eng.dso_days:.0f} / {eng.dio_days:.0f} / {eng.dpo_days:.0f} дн.",
        'Постоянные расходы': f"{eng.fixed_costs_monthly:,.0f} ₽/мес (из настроек)",
        'Переменные OPEX': f"{eng.opex_ratio*100:.1f}% выручки (из настроек)",
        'Эластичность по умолчанию': f"{eng.elast_default:.2f} (из настроек; помечается в обосновании)",
        'Порог R² для эластичности': f"{CONFIG['r2_threshold']}",
        'Сценарии спроса': f"+{(eng.opt_f-1)*100:.0f}% / −{(1-eng.pes_f)*100:.0f}% (из настроек)",
        'Макс. шаг повышения цены': f"≤{eng.price_cap*100:.0f}% (из настроек)",
        'Минимальная наценка': f"{eng.min_markup*100:.0f}% к себестоимости (из настроек)",
        'Маркетинг при росте': f"{eng.marketing_share*100:.1f}% (из настроек)",
        'Стресс-тест': f"±{eng.stress_pct*100:.0f}% (из настроек)",
        'Пороги кассовых разрывов': f"WC/поток > {eng.wc_threshold:.1f}; цикл > {eng.ccc_threshold:.0f} дн. (из настроек)",
        'Горизонт прогноза': f"{eng.horizon} дн. (из настроек)",
        'Пороги ABC': f"{CONFIG['abc_a']:.0f}% / {CONFIG['abc_b']:.0f}% выручки; XYZ: CoV {CONFIG['xyz_x']:.0f}% / {CONFIG['xyz_y']:.0f}%",
        'Порог роста портфеля (классификация)': f"{thr:.1f}% (медиана роста позиций)",
        'Цены сырья': f"из файла: {n_file}; задано вручную: {n_user}; не задано: {len(missing_raw)}",
        'Период данных': f"{dr} дн.; строк отброшено при очистке: {n_dropped}",
    }
    methodology = [
        "P&L: Выручка(без НДС) − Себестоимость сырья = Валовая прибыль; − OPEX(%) − Постоянные = EBITDA; − налог = Чистая прибыль.",
        "Точка безубыточности: постоянные расходы товара / маржинальный доход на единицу (шт/мес).",
        "Прогноз спроса: ансамбль LR+RF+GBR по лагам и календарным признакам; точность — MAPE на отложенной выборке; при нехватке данных — среднее историческое с фактическим σ.",
        "Доверительный интервал выручки: ±1.96·σ остатков ансамбля, агрегировано по позициям.",
        "Оптимум цены: формула Лернера p = c·ε/(ε−1) при ε>1.05, ограничен пользовательским шагом.",
        "Капитал на рост: Δоборотного капитала по денежному циклу (DIO+DSO−DPO) + маркетинг (настройки).",
        "Эластичность: регрессия ln(Q)~ln(P) по дневным данным; принимается при R² > 0.1, иначе значение из настроек с пометкой.",
    ]

    rep = {'total_now': round(tnow, 2), 'total_future': round(tfut, 2), 'change_pct': round(float(chg), 1),
           'profit_forecast': pf, 'raw_report': rr, 'prod_agg': pa, 'product_economics': pe,
           'target_margin': target_margin, 'daily_revenue': daily,
           'best_model': {'accuracy': round(aacc, 1) if aacc is not None else None,
                          'method': 'Ансамбль LR+RF+GBR'},
           'future_dates': fdates, 'total_forecast_revenue': frev, 'conf_interval': (cih, cih),
           'feature_importance': fi, 'actions': actions, 'increase_pct': increase_pct,
           'scenario': scenario, 'cash_warnings': cw, 'engine_params': params, 'days_range': dr,
           'period_months': pm, 'forecasts': fc, 'bcg': bcg, 'stress_test': st,
           'roi_results': [a['roi'] for a in actions], 'abc_xyz_recs': ar,
           'current_month_raw': cmr, 'raw_prices': raw_prices, 'assumptions': assumptions,
           'methodology': methodology,
           'scenario_table': scenario_comparison_table({'engine_params': params,
               'product_economics': pe, 'actions': actions})}
    return sanitize(rep)

def build_plan_report(report, mask):
    r = copy.deepcopy(report); tf = 0.0; rd = {}
    for i, (p, a) in enumerate(zip(r['profit_forecast'], r['actions'])):
        if i < len(mask) and mask[i]:
            tf += p['Прогнозная чистая прибыль']
            for c, d in a.get('raw_delta', {}).items(): rd[c] = rd.get(c, 0) + d
        else:
            tf += p['Текущая чистая прибыль']
            p['Прогнозная чистая прибыль'] = p['Текущая чистая прибыль']
            p['Изменение'] = 0.0; p['Действие'] = 'Без изменений (отклонено)'
    r['total_future'] = round(tf, 2)
    r['change_pct'] = round((tf - r['total_now']) / abs(r['total_now']) * 100, 1) if r['total_now'] else 0
    cur, rp = r.get('current_month_raw', {}), r.get('raw_prices', {})
    if cur:
        r['raw_report'] = [{'Сырьё': c, 'Текущее (мес)': round(v), 'Изменение': rd.get(c, 0),
                            'Новое': round(v + rd.get(c, 0)), 'Цена': round(rp.get(c, 0), 2)}
                           for c, v in sorted(cur.items())]
    sel = [a for a, m in zip(r['actions'], mask) if m]
    eng = FinancialEngine(r.get('engine_params', {}))
    r['stress_test'] = sensitivity_analysis(r['product_economics'], sel, r.get('period_months', 1), eng.stress_pct) \
        if sel else {'scenarios': {}, 'spread_pct': 0, 'stability': 'Нет выбранных действий'}
    r['scenario_table'] = scenario_comparison_table(r)
    r['plan_title'] = f"План: выбрано {sum(mask)} из {len(mask)} рекомендаций"
    return r

class ReportViewer(ctk.CTkToplevel):
    def __init__(self, parent, rd):
        super().__init__(parent)
        self.title(f"Отчёт — {rd.get('plan_title', 'базовый сценарий')}")
        self.geometry("1250x850"); ctk.set_appearance_mode("light")
        self.rd = rd; self.figures = []
        self.tv = ctk.CTkTabview(self, corner_radius=10)
        self.tv.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        for t in ["Допущения и методика", "KPI", "P&L по товарам", "Прогноз спроса", "Портфель",
                  "ABC-XYZ", "ROI рекомендаций", "Сценарии", "Стресс-тест", "Cash Flow", "Сырьё", "План действий"]:
            self.tv.add(t)
        self._asm(); self._kpi(); self._pl(); self._fc(); self._bcg(); self._abc(); self._roi()
        self._scn(); self._st(); self._cf(); self._raw(); self._plan()
        b = ctk.CTkFrame(self, fg_color="transparent"); b.pack(fill=tk.X, padx=10, pady=10)
        ctk.CTkButton(b, text="Экспорт PDF", command=self.export_pdf).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Экспорт Excel", command=self.export_excel).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Экспорт плана действий", command=self.export_plan, fg_color="#2ecc71").pack(side=tk.LEFT, padx=5)
    def _pf(self, tab, fig):
        c = FigureCanvasTkAgg(fig, master=tab); c.draw()
        c.get_tk_widget().pack(padx=10, pady=10, fill=tk.BOTH, expand=True); self.figures.append(fig)
    def _asm(self):
        t = ctk.CTkTextbox(self.tv.tab("Допущения и методика"), font=ctk.CTkFont(size=12))
        t.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        t.insert("0.0", "ДОПУЩЕНИЯ МОДЕЛИ (все входы — из файла данных или из настроек пользователя)\n\n")
        for k, v in self.rd.get('assumptions', {}).items(): t.insert("end", f"{k}: {v}\n")
        t.insert("end", "\nИсточники эластичности по позициям:\n")
        for p in self.rd['profit_forecast']:
            t.insert("end", f"• {p['Товар']}: {p['Источник эластичности']}\n")
        t.insert("end", "\nМЕТОДИКА РАСЧЁТА\n\n")
        for m in self.rd.get('methodology', []): t.insert("end", f"— {m}\n")
    def _kpi(self):
        d = self.rd; fig = plt.Figure(figsize=(10, 4), dpi=100, facecolor='white')
        ax = fig.add_subplot(111); ax.set_facecolor('#f8f9fa')
        ax.plot(d["daily_revenue"]["Дата"], d["daily_revenue"]["Выручка"], 'b-', lw=2, label="Выручка (факт)")
        ax.plot(d["future_dates"], d["total_forecast_revenue"], 'r--', lw=2, label="Прогноз")
        ax.fill_between(d["future_dates"], np.array(d["total_forecast_revenue"]) - d["conf_interval"][0],
                        np.array(d["total_forecast_revenue"]) + d["conf_interval"][1], color='red', alpha=0.15)
        ax.set_title("Динамика выручки и прогноз (доверительный интервал ±1.96σ)")
        ax.grid(True, ls='--', alpha=0.5); ax.legend()
        self._pf(self.tv.tab("KPI"), fig)
    def _pl(self):
        t = ctk.CTkTextbox(self.tv.tab("P&L по товарам"), font=ctk.CTkFont(size=11))
        t.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        t.insert("0.0", "P&L ПО ТОВАРАМ (чистая прибыль после налога)\n\n")
        t.insert("end", pd.DataFrame(self.rd['profit_forecast']).to_string(index=False))
    def _fc(self):
        tab = self.tv.tab("Прогноз спроса"); fc = self.rd.get('forecasts', {})
        if not fc: ctk.CTkLabel(tab, text="Нет данных").pack(); return
        cols = 2; rows = (len(fc) + 1) // 2
        fig = plt.Figure(figsize=(12, 3 * rows), dpi=100, facecolor='white')
        for i, (p, f) in enumerate(fc.items()):
            ax = fig.add_subplot(rows, cols, i + 1); h = f['historical'].tail(60)
            ax.plot(h['Дата'], h['Количество'], 'b-', lw=1.5, label="Факт")
            ax.plot(f['forecast_dates'], f['forecast_units'], 'r--', lw=2,
                    label=f"Прогноз (точность {f['accuracy']:.0f}%)" if f['accuracy'] is not None else "Прогноз (среднее)")
            ax.set_title(p, fontsize=11); ax.grid(alpha=0.4, ls='--'); ax.legend(fontsize=8)
        fig.tight_layout(); self._pf(tab, fig)
        t = ctk.CTkTextbox(tab, height=180); t.pack(fill=tk.X, padx=10, pady=10)
        t.insert("0.0", "ЧИСЛОВОЙ ПРОГНОЗ СПРОСА (шт/день, среднее по горизонту ±1.96σ остатков)\n\n")
        for p, f in fc.items():
            acc = f"{f['accuracy']:.1f}%" if f['accuracy'] is not None else "н/д (мало данных)"
            t.insert("end", f"{p}: {np.mean(f['forecast_units']):.1f} ± {1.96*f['forecast_std']:.1f} шт/день | "
                            f"точность {acc} | метод: {f['method']}\n")
    def _bcg(self):
        tab = self.tv.tab("Портфель"); bcg = self.rd.get('bcg', {})
        if not bcg: ctk.CTkLabel(tab, text="Нет данных").pack(); return
        fig = plt.Figure(figsize=(9, 6), dpi=100, facecolor='white'); ax = fig.add_subplot(111)
        col = {'Лидер (рост)': '#2ecc71', 'Лидер (стабильность)': '#3498db',
               'Перспективный': '#f39c12', 'Низкий приоритет': '#e74c3c'}
        for p, i in bcg.items():
            ax.scatter(i['share_pct'], i['growth_pct'], s=200, c=col.get(i['quadrant'], '#95a5a6'),
                       alpha=0.7, edgecolors='black')
            ax.annotate(p, (i['share_pct'], i['growth_pct']), textcoords="offset points", xytext=(5, 5), fontsize=9)
        ax.axhline(y=i['threshold_pct'], color='gray', ls='--', alpha=0.5)
        ax.axvline(x=15, color='gray', ls='--', alpha=0.5)
        ax.set_xlabel("Доля в выручке, %"); ax.set_ylabel("Прогнозный рост спроса, %")
        ax.set_title("Классификация портфеля (доля x рост)"); ax.grid(alpha=0.3, ls='--')
        self._pf(tab, fig)
        t = ctk.CTkTextbox(tab, height=160); t.pack(fill=tk.X, padx=10, pady=10)
        t.insert("0.0", "СТРАТЕГИИ ПО ПОЗИЦИЯМ\n\n")
        for p, i in bcg.items():
            t.insert("end", f"{p}: {i['quadrant']} | доля {i['share_pct']}%, рост {i['growth_pct']}%\n  {i['strategy']}\n\n")
    def _abc(self):
        tab = self.tv.tab("ABC-XYZ"); pa = self.rd['prod_agg']
        if pa.empty or 'abc_class' not in pa.columns: ctk.CTkLabel(tab, text="Нет данных").pack(); return
        fig = plt.Figure(figsize=(10, 5), dpi=100, facecolor='white')
        ax = fig.add_subplot(121); ax.pie(pa['total_revenue'], labels=pa['Товар'], autopct='%1.1f%%')
        ax.set_title("Структура выручки (ABC)")
        ax2 = fig.add_subplot(122); cl = pa['abc_class'].value_counts()
        ax2.bar(cl.index, cl.values); ax2.set_title("Распределение по классам"); fig.tight_layout()
        self._pf(tab, fig)
        t = ctk.CTkTextbox(tab, height=200); t.pack(fill=tk.BOTH, padx=10, pady=10)
        t.insert("0.0", "РЕКОМЕНДАЦИИ ПО УПРАВЛЕНИЮ ЗАПАСАМИ\n\n")
        for r in self.rd.get('abc_xyz_recs', []):
            t.insert("end", f"{r['Класс']} — {r['Товар']}\n  {r['Рекомендация по запасам']}\n\n")
    def _roi(self):
        tab = self.tv.tab("ROI рекомендаций"); d = self.rd
        fig = plt.Figure(figsize=(10, 5), dpi=100, facecolor='white'); ax = fig.add_subplot(111)
        ef = [a['roi']['delta_profit_month'] for a in d['actions']]
        ax.bar([a['Товар'] for a in d['actions']], ef, color=['#2ecc71' if r > 0 else '#e74c3c' for r in ef])
        ax.axhline(y=0, color='black', lw=0.8); ax.set_title("Эффект рекомендаций, ₽/мес")
        ax.grid(axis='y', alpha=0.3, ls='--'); self._pf(tab, fig)
        t = ctk.CTkTextbox(tab, height=200); t.pack(fill=tk.X, padx=10, pady=10)
        t.insert("0.0", "ДЕТАЛИЗАЦИЯ ROI\n\n")
        for a in d['actions']:
            r = a['roi']
            t.insert("end", f"{a['Тип']} — {a['Товар']}\n  Эффект: {r['delta_profit_month']:+,.0f} ₽/мес | "
                            f"Капитал: {r['capital_required']:,.0f} ₽ | Окупаемость: {r['payback_months']} мес | "
                            f"Риск: {r['risk_level']}\n\n")
    def _scn(self):
        tab = self.tv.tab("Сценарии"); rows = self.rd.get('scenario_table', [])
        if not rows: ctk.CTkLabel(tab, text="Нет данных").pack(); return
        fig = plt.Figure(figsize=(9, 5), dpi=100, facecolor='white'); ax = fig.add_subplot(111)
        x = np.arange(len(rows)); w = 0.35
        ax.bar(x - w/2, [r['Без изменений'] for r in rows], w, label="Без изменений", color='#95a5a6')
        ax.bar(x + w/2, [r['С планом'] for r in rows], w, label="С планом", color='#2ecc71')
        ax.set_xticks(x); ax.set_xticklabels([r['Сценарий'] for r in rows])
        ax.set_ylabel("Чистая прибыль за период, ₽"); ax.set_title("Сравнение сценариев спроса")
        ax.legend(); ax.grid(axis='y', alpha=0.3, ls='--'); self._pf(tab, fig)
        t = ctk.CTkTextbox(tab, height=140); t.pack(fill=tk.X, padx=10, pady=10)
        t.insert("0.0", pd.DataFrame(rows).to_string(index=False))
    def _st(self):
        tab = self.tv.tab("Стресс-тест"); st = self.rd.get('stress_test', {})
        if not st or not st.get('scenarios'): ctk.CTkLabel(tab, text="Нет выбранных действий").pack(); return
        fig = plt.Figure(figsize=(8, 5), dpi=100, facecolor='white'); ax = fig.add_subplot(111)
        ax.bar(list(st['scenarios'].keys()), list(st['scenarios'].values()), color=['#3498db', '#e74c3c', '#2ecc71'])
        ax.set_title(f"Прибыль/мес при отклонении спроса (разброс {st['spread_pct']}%)")
        ax.grid(axis='y', alpha=0.3, ls='--'); self._pf(tab, fig)
        t = ctk.CTkTextbox(tab, height=100); t.pack(fill=tk.X, padx=10, pady=10)
        t.insert("0.0", f"Устойчивость плана: {st['stability']}\n")
    def _cf(self):
        tab = self.tv.tab("Cash Flow"); d = self.rd
        fr = ctk.CTkFrame(tab, fg_color="#f8f9fa", corner_radius=8); fr.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        ctk.CTkLabel(fr, text="Денежный поток и оборотный капитал", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        e = FinancialEngine(d.get('engine_params', {})); pm = d.get('period_months', 1)
        mr = d['prod_agg']['total_revenue'].sum() / pm; mc = d['prod_agg']['total_cogs'].sum() / pm
        wc = e.working_capital(mr, mc)
        eb = mr - mc - mr * e.opex_ratio - e.fixed_costs_monthly
        for l, v in [("Выручка, ₽/мес", f"{mr:,.0f}"), ("EBITDA, ₽/мес", f"{eb:,.0f}"),
                     ("Потребность в оборотном капитале, ₽", f"{wc['working_capital']:,.0f}"),
                     ("Денежный цикл, дн.", f"{wc['cash_conversion_cycle']:.0f}")]:
            r = ctk.CTkFrame(fr, fg_color="transparent"); r.pack(fill=tk.X, pady=2)
            ctk.CTkLabel(r, text=l, font=ctk.CTkFont(size=13)).pack(side=tk.LEFT, padx=20)
            ctk.CTkLabel(r, text=v, font=ctk.CTkFont(size=13, weight="bold"), text_color="#0057b3").pack(side=tk.RIGHT, padx=20)
        if d.get('cash_warnings'):
            wf = ctk.CTkFrame(fr, fg_color="#fff3cd", corner_radius=8); wf.pack(fill=tk.X, padx=20, pady=10)
            ctk.CTkLabel(wf, text="Выявленные риски:", font=ctk.CTkFont(size=13, weight="bold"),
                         text_color="#856404").pack(pady=5)
            for w in d['cash_warnings']:
                ctk.CTkLabel(wf, text=f"— {w['msg']}", font=ctk.CTkFont(size=11), text_color="#856404",
                             wraplength=800, justify="left").pack(anchor="w", padx=10, pady=2)
    def _raw(self):
        tab = self.tv.tab("Сырьё")
        if not self.rd['raw_report']: ctk.CTkLabel(tab, text="Нет данных").pack(); return
        t = ctk.CTkTextbox(tab, font=ctk.CTkFont(size=11)); t.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        t.insert("0.0", "ПЛАН ЗАКУПОК СЫРЬЯ (мес)\n\n")
        t.insert("end", pd.DataFrame(self.rd['raw_report']).to_string(index=False))
    def _plan(self):
        t = ctk.CTkTextbox(self.tv.tab("План действий"), font=ctk.CTkFont(size=11))
        t.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        t.insert("0.0", "ПРИОРИТЕТНЫЙ ПЛАН (по эффекту, ₽/мес)\n\n")
        for i, a in enumerate(sorted(self.rd['actions'], key=lambda x: x['roi']['delta_profit_month'], reverse=True), 1):
            r = a['roi']
            t.insert("end", f"{i}. {a['Тип']} — {a['Товар']}\n   {a['Обоснование']}\n"
                            f"   Эффект: {r['delta_profit_month']:+,.0f} ₽/мес | Капитал: "
                            f"{r['capital_required']:,.0f} ₽ | Окупаемость: {r['payback_months']} мес\n\n")
    def export_pdf(self):
        p = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not p: return
        try:
            with PdfPages(p) as pdf:
                for f in self.figures: pdf.savefig(f, bbox_inches='tight')
            messagebox.showinfo("Успех", f"PDF сохранён: {p}")
        except Exception as e: messagebox.showerror("Ошибка", str(e))
    def export_excel(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not p: return
        try:
            d = self.rd
            with pd.ExcelWriter(p, engine='openpyxl') as w:
                pd.DataFrame([f"{k}: {v}" for k, v in d.get('assumptions', {}).items()],
                             columns=['Допущение']).to_excel(w, sheet_name='Допущения', index=False)
                pd.DataFrame(d['methodology'], columns=['Методика']).to_excel(w, sheet_name='Методика', index=False)
                pd.DataFrame(d['profit_forecast']).to_excel(w, sheet_name='P&L', index=False)
                if d['raw_report']: pd.DataFrame(d['raw_report']).to_excel(w, sheet_name='Сырьё', index=False)
                d['prod_agg'].to_excel(w, sheet_name='ABC-XYZ', index=False)
                pd.DataFrame([{'Товар': k, **v} for k, v in d.get('bcg', {}).items()]).to_excel(w, sheet_name='Портфель', index=False)
                pd.DataFrame(d.get('scenario_table', [])).to_excel(w, sheet_name='Сценарии', index=False)
                pd.DataFrame(d['roi_results']).to_excel(w, sheet_name='ROI', index=False)
            messagebox.showinfo("Успех", f"Excel сохранён: {p}")
        except Exception as e: messagebox.showerror("Ошибка", str(e))
    def export_plan(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not p: return
        try:
            d = self.rd
            df = pd.DataFrame([{**a, **a['roi']} for a in d['actions']])
            df = df.sort_values('delta_profit_month', ascending=False)
            df['Приоритет'] = ['Высокий' if x > 0 else 'Не реализовывать' for x in df['delta_profit_month']]
            with pd.ExcelWriter(p, engine='openpyxl') as w:
                df.to_excel(w, sheet_name='План действий', index=False)
            messagebox.showinfo("Успех", f"План сохранён: {p}")
        except Exception as e: messagebox.showerror("Ошибка", str(e))

class FinScopeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(CONFIG["app_name"]); self.geometry("1300x800"); self.minsize(1100, 700)
        ctk.set_appearance_mode("light"); ctk.set_default_color_theme("blue")
        self.file_path = ctk.StringVar()
        self.target_margin = ctk.StringVar(value="20"); self.increase_pct = ctk.StringVar(value="20")
        self.scenario = StringVar(value="realistic")
        self.report_data = None; self.db = HistoryDB()
        self._running = False; self._stop_flag = False; self.import_result = None; self.stop_btn = None
        self.fixed_costs = DoubleVar(value=500000); self.opex_ratio = DoubleVar(value=15)
        self.dso_days = IntVar(value=15); self.dio_days = IntVar(value=30); self.dpo_days = IntVar(value=20)
        self.vat_rate = DoubleVar(value=20); self.profit_tax_rate = DoubleVar(value=20)
        self.elast_default = DoubleVar(value=1.0); self.price_cap = DoubleVar(value=15)
        self.marketing_pct = DoubleVar(value=2); self.price_change_pct = DoubleVar(value=0.5)
        self.stress_pct = DoubleVar(value=20); self.min_markup = DoubleVar(value=5)
        self.wc_threshold = DoubleVar(value=2.0); self.ccc_threshold = DoubleVar(value=45)
        self.horizon_days = IntVar(value=7)
        self.opt_pct = StringVar(value="15"); self.pes_pct = StringVar(value="15")
        self._build_ui(); self._load_settings(); self.show_dashboard()
        ok, msg = LicenseManager.check_license()
        if not ok:
            messagebox.showerror("Лицензия", f"{msg}\n\nПриобретите лицензию на сайте."); self.quit()
        else: self.progress_var.set(msg)
        if not os.path.exists("onboarding_done.txt"): self._onboarding()

    def report_callback_exception(self, exc, val, tb):
        logger.error(f"UI callback: {exc}")
    def _num(self, var, default, cast=int):
        try: return cast(float(var.get()))
        except Exception: return default
    def _params(self):
        return {'fixed_costs': self.fixed_costs.get(), 'opex_ratio': self.opex_ratio.get() / 100,
                'dso_days': self.dso_days.get(), 'dio_days': self.dio_days.get(), 'dpo_days': self.dpo_days.get(),
                'vat_rate': self.vat_rate.get() / 100, 'profit_tax_rate': self.profit_tax_rate.get() / 100,
                'elast_default': self.elast_default.get(), 'price_cap_pct': self.price_cap.get(),
                'marketing_pct': self.marketing_pct.get(), 'price_change_pct': self.price_change_pct.get(),
                'stress_pct': self.stress_pct.get(), 'min_markup_pct': self.min_markup.get(),
                'wc_threshold': self.wc_threshold.get(), 'ccc_threshold': self.ccc_threshold.get(),
                'horizon_days': self.horizon_days.get(),
                'opt_pct': self._num(self.opt_pct, 15), 'pes_pct': self._num(self.pes_pct, 15),
                'user_raw_prices': load_user_raw_prices()}

    def _build_ui(self):
        self.sidebar = ctk.CTkFrame(self, width=230, corner_radius=0, fg_color="#f0f0f0")
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y); self.sidebar.pack_propagate(False)
        ctk.CTkLabel(self.sidebar, text="FinScope Pro", font=ctk.CTkFont(size=18, weight="bold"),
                     text_color="#0057b3").pack(pady=(20, 5))
        ctk.CTkLabel(self.sidebar, text=f"v{CONFIG['version']}", font=ctk.CTkFont(size=10), text_color="gray").pack()
        for t, c in [("Дашборд", self.show_dashboard), ("Загрузка данных", self.show_upload),
                     ("Цены сырья", self.show_raw_prices), ("Настройки и допущения", self.show_settings),
                     ("Анализ чувствительности", self.show_whatif), ("Рекомендации", self.show_recommendations),
                     ("Сценарии", self.show_scenarios), ("История", self.show_history),
                     ("О программе", self.show_about), ("Помощь", self.show_help)]:
            ctk.CTkButton(self.sidebar, text=t, command=c, fg_color="transparent", text_color="#333",
                          hover_color="#d0d0d0", anchor="w", font=ctk.CTkFont(size=13)).pack(fill=tk.X, padx=10, pady=2)
        self.main_frame = ctk.CTkFrame(self, corner_radius=10, fg_color="white")
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=15, pady=15)
        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=(0, 10))
        self.progress_var = tk.StringVar(value="Готов")
        ctk.CTkLabel(self.progress_frame, textvariable=self.progress_var, font=ctk.CTkFont(size=12)).pack(side=tk.LEFT, padx=5)
        self.progress = ctk.CTkProgressBar(self.progress_frame, height=10, width=300)
        self.progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True); self.progress.set(0)

    def log_message(self, m): self.progress_var.set(m)
    def _sp(self, v):
        try:
            if self.progress.winfo_exists(): self.progress.set(v)
        except Exception: pass
    def update_progress(self, m, v): self.progress_var.set(m); self._sp(v / 100)
    def clear_main(self):
        for w in self.main_frame.winfo_children(): w.destroy()

    def _collect_materials(self):
        try:
            di = self.import_result if self.import_result is not None else self.file_path.get()
            if not di: return set(), {}
            df = load_production_data(di)
            df = df.rename(columns={o: t for t, o in detect_columns(df).items()})
            if 'Компоненты' not in df.columns: return set(), {}
            mats = set()
            for _, r in df.iterrows():
                mats.update(parse_components_safe(r.get('Компоненты', '')).keys())
            known = load_raw_prices(di); known.update(load_user_raw_prices())
            return mats, known
        except Exception:
            return set(), {}

    def show_raw_prices(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="Цены сырья (ручной ввод)", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        ctk.CTkLabel(self.main_frame, text="Программа не использует цены по умолчанию: всё, что не найдено "
                                           "в файле, заполняется только здесь.", font=ctk.CTkFont(size=12)).pack(pady=5)
        mats, known = self._collect_materials()
        if not mats:
            ctk.CTkLabel(self.main_frame, text="Компоненты в данных не найдены. Загрузите файл с колонкой "
                                               "«Компоненты» (формат: 'Сырьё_A 2 Сырьё_B 3').").pack(pady=20)
            return
        fr = ctk.CTkScrollableFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8)
        fr.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.rp_vars = {}
        for i, m in enumerate(sorted(mats)):
            ctk.CTkLabel(fr, text=m, font=ctk.CTkFont(size=12), width=220, anchor="w").grid(row=i, column=0, padx=8, pady=3)
            v = StringVar(value=str(known.get(m, "")))
            ctk.CTkEntry(fr, textvariable=v, width=120, placeholder_text="₽ за ед.").grid(row=i, column=1, padx=8, pady=3)
            self.rp_vars[m] = v
            if m not in known:
                ctk.CTkLabel(fr, text="не задано", text_color="#e67e22", font=ctk.CTkFont(size=10)).grid(row=i, column=2, padx=5)
        b = ctk.CTkFrame(self.main_frame, fg_color="transparent"); b.pack(pady=10)
        ctk.CTkButton(b, text="Сохранить цены", command=self.save_raw_prices, fg_color="#0057b3").pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Пересчитать отчёт", command=self.start_generation).pack(side=tk.LEFT, padx=5)

    def save_raw_prices(self):
        try:
            out = {}
            for m, v in self.rp_vars.items():
                s = v.get().strip().replace(',', '.')
                if s: out[m] = float(s)
            with open('raw_prices_user.json', 'w', encoding='utf-8') as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Успех", f"Сохранено цен сырья: {len(out)}.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def show_dashboard(self):
        self.clear_main()
        if not self.report_data:
            ctk.CTkLabel(self.main_frame, text="Дашборд", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
            ctk.CTkLabel(self.main_frame, text="Загрузите данные или сгенерируйте отчёт.", font=ctk.CTkFont(size=14)).pack(pady=20)
            ctk.CTkButton(self.main_frame, text="Загрузить данные", command=self.show_upload, fg_color="#0057b3").pack(pady=10)
            ctk.CTkButton(self.main_frame, text="Демо-пример", command=self.load_demo).pack(pady=5)
            return
        d = self.report_data
        k = ctk.CTkFrame(self.main_frame, fg_color="transparent"); k.pack(fill=tk.X, pady=10)
        k.grid_columnconfigure((0, 1, 2, 3), weight=1)
        acc = d['best_model']['accuracy']
        for i, (l, v, c) in enumerate([
            ("Чистая прибыль (факт)", f"{d['total_now']:,.0f} ₽", "#34495e"),
            ("Чистая прибыль (план)", f"{d['total_future']:,.0f} ₽", "#0057b3"),
            ("Изменение", f"{d['change_pct']:+.1f}%", "#2ecc71" if d['change_pct'] > 0 else "#e74c3c"),
            ("Точность прогноза спроса", f"{acc:.1f}%" if acc is not None else "н/д", "#27ae60")]):
            f = ctk.CTkFrame(k, fg_color="#f8f9fa", corner_radius=8); f.grid(row=0, column=i, padx=8, pady=5, sticky="nsew")
            ctk.CTkLabel(f, text=l, font=ctk.CTkFont(size=11), text_color="#6c757d").pack(pady=(5, 0))
            ctk.CTkLabel(f, text=v, font=ctk.CTkFont(size=17, weight="bold"), text_color=c).pack(pady=(0, 5))
        if d.get('cash_warnings'):
            wf = ctk.CTkFrame(self.main_frame, fg_color="#fff3cd", corner_radius=8); wf.pack(fill=tk.X, padx=10, pady=5)
            for w in d['cash_warnings']:
                ctk.CTkLabel(wf, text=f"Риск: {w['msg']}", font=ctk.CTkFont(size=11), text_color="#856404",
                             wraplength=1000, justify="left").pack(anchor="w", padx=10, pady=3)
        pa = d['prod_agg']
        if not pa.empty:
            ctk.CTkLabel(self.main_frame, text="Показатели по товарам", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=10, pady=5)
            pf = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8); pf.pack(fill=tk.X, padx=10, pady=5)
            for j, h in enumerate(["Товар", "ABC-XYZ", "Класс", "Маржа %", "Прибыль ₽", "Безубыточность"]):
                ctk.CTkLabel(pf, text=h, font=ctk.CTkFont(size=11, weight="bold")).grid(row=0, column=j, padx=8, pady=2)
            for i, r in pa.head(10).iterrows():
                e = d['product_economics'].get(r['Товар'], {})
                nm = e.get('pl', {}).get('net_margin_pct', 0)
                c = "#2ecc71" if nm >= 15 else "#f39c12" if nm > 0 else "#e74c3c"
                vals = [r['Товар'], r.get('abc_xyz', '??'), d.get('bcg', {}).get(r['Товар'], {}).get('quadrant', ''),
                        f"{nm:.1f}%", f"{e.get('pl', {}).get('net_profit', 0):,.0f}",
                        "Да" if e.get('break_even', {}).get('is_profitable') else "Нет"]
                for j, v in enumerate(vals):
                    ctk.CTkLabel(pf, text=v, font=ctk.CTkFont(size=11), text_color=c if j in (0, 3) else None).grid(row=i+1, column=j, padx=8, pady=2)
        fig = plt.Figure(figsize=(10, 3.5), dpi=100, facecolor='white')
        ax = fig.add_subplot(111); ax.set_facecolor('#f8f9fa')
        ax.plot(d["daily_revenue"]["Дата"], d["daily_revenue"]["Выручка"], 'b-', lw=2, label="Выручка (факт)")
        ax.plot(d["future_dates"], d["total_forecast_revenue"], 'r--', lw=2, label="Прогноз")
        ax.fill_between(d["future_dates"], np.array(d["total_forecast_revenue"]) - d["conf_interval"][0],
                        np.array(d["total_forecast_revenue"]) + d["conf_interval"][1], color='red', alpha=0.15)
        ax.set_title("Динамика выручки (доверительный интервал ±1.96σ)"); ax.grid(True, ls='--', alpha=0.5); ax.legend()
        c = FigureCanvasTkAgg(fig, master=self.main_frame); c.draw()
        c.get_tk_widget().pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        sf = ctk.CTkFrame(self.main_frame, fg_color="transparent"); sf.pack(fill=tk.X, pady=5)
        ctk.CTkLabel(sf, text="Сценарий:", font=ctk.CTkFont(size=12)).pack(side=tk.LEFT, padx=10)
        for s, l in [("optimistic", "Оптимистичный"), ("realistic", "Реалистичный"), ("pessimistic", "Пессимистичный")]:
            ctk.CTkRadioButton(sf, text=l, variable=self.scenario, value=s, command=self.update_scenario).pack(side=tk.LEFT, padx=10)
        ctk.CTkButton(self.main_frame, text="Открыть полный отчёт", command=self.show_full_report,
                      fg_color="#0057b3", font=ctk.CTkFont(size=13)).pack(pady=10)

    def show_whatif(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="Анализ чувствительности: постоянные расходы и OPEX",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        if not self.report_data:
            ctk.CTkLabel(self.main_frame, text="Сначала сгенерируйте отчёт.").pack(pady=20); return
        base_f = float(self.report_data['engine_params'].get('fixed_costs', 500000))
        base_o = float(self.report_data['engine_params'].get('opex_ratio', 0.15)) * 100
        self.wi_fixed = DoubleVar(value=base_f); self.wi_opex = DoubleVar(value=base_o)
        fr = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8); fr.pack(fill=tk.X, padx=20, pady=10)
        ctk.CTkLabel(fr, text="Постоянные расходы, ₽/мес:").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        ctk.CTkSlider(fr, from_=0, to=base_f * 3, variable=self.wi_fixed, command=lambda v: self._whatif()).grid(row=0, column=1, padx=10, pady=8, sticky="ew")
        self.wi_f_lab = ctk.CTkLabel(fr, text=""); self.wi_f_lab.grid(row=0, column=2, padx=10)
        ctk.CTkLabel(fr, text="OPEX, % выручки:").grid(row=1, column=0, padx=10, pady=8, sticky="w")
        ctk.CTkSlider(fr, from_=0, to=40, variable=self.wi_opex, command=lambda v: self._whatif()).grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        self.wi_o_lab = ctk.CTkLabel(fr, text=""); self.wi_o_lab.grid(row=1, column=2, padx=10)
        self.wi_res = ctk.CTkTextbox(self.main_frame, font=ctk.CTkFont(size=12))
        self.wi_res.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self._whatif()

    def _whatif(self):
        if not self.report_data: return
        rep = self.report_data
        fixed = self.wi_fixed.get(); opex = self.wi_opex.get() / 100
        self.wi_f_lab.configure(text=f"{fixed:,.0f} ₽"); self.wi_o_lab.configure(text=f"{opex*100:.1f}%")
        tax = float(rep['engine_params'].get('profit_tax_rate', 0.2)); pm = rep.get('period_months', 1)
        shares = sum(e['revenue_share'] for e in rep['product_economics'].values()) or 1
        total_m, below, lines = 0.0, [], []
        for p, e in rep['product_economics'].items():
            pn = e['unit_economics']['price_net_vat']; cu = e['cogs_per_unit']
            cm = pn - cu - pn * opex
            am = e['units_period'] / pm
            fm = fixed * (e['revenue_share'] / shares)
            if cm > 0:
                be = fm / cm; ok = am > be
            else:
                be = math.inf; ok = False
            total_m += (cm * am - fm) * (1 - tax)
            if not ok: below.append(p)
            lines.append(f"{'[+] ' if ok else '[-] '}{p}: маржинальный доход {cm:,.0f} ₽/шт; "
                         f"БЕ {be:,.0f} шт/мес (факт {am:,.0f})" if math.isfinite(be)
                         else f"[-] {p}: маржинальный доход <= 0")
        self.wi_res.delete("0.0", tk.END)
        self.wi_res.insert("0.0", f"ИТОГО чистая прибыль: {total_m:,.0f} ₽/мес\nПозиций ниже точки безубыточности: "
                                  f"{len(below)}{f' ({', '.join(below)})' if below else ''}\n\n" + "\n".join(lines))

    def show_upload(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="Загрузка данных", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        ff = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8); ff.pack(fill=tk.X, padx=20, pady=10)
        ctk.CTkLabel(ff, text="Файл:", font=ctk.CTkFont(size=13)).grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkEntry(ff, textvariable=self.file_path, width=400).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkButton(ff, text="Обзор", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(ff, text="Мастер импорта", command=self.open_wizard).grid(row=0, column=3, padx=5, pady=5)
        ctk.CTkButton(ff, text="Демо-пример", command=self.load_demo).grid(row=0, column=4, padx=5, pady=5)
        pf = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8); pf.pack(fill=tk.X, padx=20, pady=10)
        g = ctk.CTkFrame(pf, fg_color="transparent"); g.pack(padx=10, pady=5)
        ctk.CTkLabel(g, text="Целевая чистая маржа, %:").grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkEntry(g, textvariable=self.target_margin, width=80).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkLabel(g, text="План роста выпуска, %:").grid(row=1, column=0, padx=5, pady=5)
        ctk.CTkEntry(g, textvariable=self.increase_pct, width=80).grid(row=1, column=1, padx=5, pady=5)
        ctk.CTkButton(self.main_frame, text="Сгенерировать отчёт", command=self.start_generation,
                      fg_color="#0057b3", font=ctk.CTkFont(size=14, weight="bold"), height=40).pack(pady=20)

    def show_settings(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="Настройки и допущения модели", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        fr = ctk.CTkScrollableFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8)
        fr.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        rows = [("Постоянные расходы, ₽/мес", self.fixed_costs, 1), ("OPEX, % выручки", self.opex_ratio, 1),
                ("DSO, дн.", self.dso_days, 1), ("DIO, дн.", self.dio_days, 1), ("DPO, дн.", self.dpo_days, 1),
                ("НДС, %", self.vat_rate, 1), ("Налог на прибыль, %", self.profit_tax_rate, 1),
                ("Эластичность по умолчанию", self.elast_default, 1), ("Макс. шаг цены, %", self.price_cap, 1),
                ("Мин. наценка к себестоимости, %", self.min_markup, 1),
                ("Маркетинг при росте, %", self.marketing_pct, 1), ("Затраты на смену цены, %", self.price_change_pct, 1),
                ("Стресс-тест, ±%", self.stress_pct, 1), ("Порог WC/денежный поток", self.wc_threshold, 1),
                ("Порог денежного цикла, дн.", self.ccc_threshold, 1),
                ("Горизонт прогноза, дн.", self.horizon_days, 1)]
        for i, (l, rv, m) in enumerate(rows):
            ctk.CTkLabel(fr, text=l, font=ctk.CTkFont(size=12)).grid(row=i, column=0, padx=10, pady=4, sticky="w")
            disp = StringVar(value=str(rv.get() * m))
            ctk.CTkEntry(fr, textvariable=disp, width=100).grid(row=i, column=1, padx=5, pady=4)
            def mk(rv=rv, dv=disp, m=m):
                def sy(*a):
                    try: rv.set(float(dv.get()) / m)
                    except Exception: pass
                return sy
            disp.trace_add("write", mk())
        b = ctk.CTkFrame(self.main_frame, fg_color="transparent"); b.pack(pady=10)
        ctk.CTkButton(b, text="Сохранить", command=self.save_settings, fg_color="#0057b3").pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Пересчитать отчёт", command=self.start_generation).pack(side=tk.LEFT, padx=5)

    def save_settings(self):
        try:
            with open('financial_settings.json', 'w', encoding='utf-8') as f:
                json.dump({'fixed_costs': self.fixed_costs.get(), 'opex_ratio': self.opex_ratio.get() / 100,
                           'dso_days': self.dso_days.get(), 'dio_days': self.dio_days.get(),
                           'dpo_days': self.dpo_days.get(), 'vat_rate': self.vat_rate.get() / 100,
                           'profit_tax_rate': self.profit_tax_rate.get() / 100,
                           'elast_default': self.elast_default.get(), 'price_cap_pct': self.price_cap.get(),
                           'min_markup_pct': self.min_markup.get(), 'marketing_pct': self.marketing_pct.get(),
                           'price_change_pct': self.price_change_pct.get(), 'stress_pct': self.stress_pct.get(),
                           'wc_threshold': self.wc_threshold.get(), 'ccc_threshold': self.ccc_threshold.get(),
                           'horizon_days': self.horizon_days.get(),
                           'opt_pct': self._num(self.opt_pct, 15), 'pes_pct': self._num(self.pes_pct, 15)}, f, indent=2)
            messagebox.showinfo("Успех", "Настройки сохранены.")
        except Exception as e: messagebox.showerror("Ошибка", str(e))

    def _load_settings(self):
        if os.path.exists('financial_settings.json'):
            try:
                with open('financial_settings.json', 'r', encoding='utf-8') as f: s = json.load(f)
                mp = {'fixed_costs': self.fixed_costs, 'dso_days': self.dso_days, 'dio_days': self.dio_days,
                      'dpo_days': self.dpo_days, 'elast_default': self.elast_default, 'price_cap_pct': self.price_cap,
                      'min_markup_pct': self.min_markup, 'marketing_pct': self.marketing_pct,
                      'price_change_pct': self.price_change_pct, 'stress_pct': self.stress_pct,
                      'wc_threshold': self.wc_threshold, 'ccc_threshold': self.ccc_threshold,
                      'horizon_days': self.horizon_days}
                for k, v in mp.items():
                    if k in s: v.set(s[k])
                if 'opex_ratio' in s: self.opex_ratio.set(s['opex_ratio'] * 100)
                if 'vat_rate' in s: self.vat_rate.set(s['vat_rate'] * 100)
                if 'profit_tax_rate' in s: self.profit_tax_rate.set(s['profit_tax_rate'] * 100)
                if 'opt_pct' in s: self.opt_pct.set(str(s['opt_pct']))
                if 'pes_pct' in s: self.pes_pct.set(str(s['pes_pct']))
            except Exception: pass

    def browse_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("JSON", "*.json")])
        if p: self.file_path.set(p)
    def open_wizard(self):
        w = ImportWizard(self); self.wait_window(w)
        if self.import_result is not None: self.file_path.set("Импортировано через мастер")
    def load_demo(self):
        self.file_path.set(generate_demo_data())
        messagebox.showinfo("Успех", "Демо-данные загружены. Нажмите «Сгенерировать отчёт».")

    def start_generation(self):
        if self._running: return
        if not self.file_path.get() and self.import_result is None:
            messagebox.showerror("Ошибка", "Загрузите файл или используйте демо-пример."); return
        self._running = True; self._stop_flag = False; self._sp(0)
        self.stop_btn = ctk.CTkButton(self.progress_frame, text="Остановить", command=self.stop_generation, fg_color="#e74c3c")
        self.stop_btn.pack(side=tk.RIGHT, padx=5)
        threading.Thread(target=self._gen, daemon=True).start()
    def stop_generation(self):
        self._stop_flag = True; self.progress_var.set("Остановка...")
        if self.stop_btn and self.stop_btn.winfo_exists(): self.stop_btn.destroy()

    def _gen(self):
        try:
            di = self.import_result if self.import_result is not None else self.file_path.get()
            rep = generate_report_data(di, self._num(self.target_margin, 20), self._num(self.increase_pct, 20),
                                       self.scenario.get(), self.update_progress, lambda: self._stop_flag,
                                       params=self._params())
            if self._stop_flag: self.progress_var.set("Остановлен"); return
            if rep:
                self.report_data = rep
                self.after(0, self.show_dashboard)
                self.after(0, lambda: messagebox.showinfo("Успех", "Отчёт сгенерирован."))
                self.db.save_project(f"Проект от {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", rep,
                                     scenario=rep['scenario'],
                                     kpi_snapshot={'net_profit_now': rep['total_now'],
                                                   'net_profit_forecast': rep['total_future'],
                                                   'change_pct': rep['change_pct'],
                                                   'ml_accuracy': rep['best_model']['accuracy'],
                                                   'stability': rep.get('stress_test', {}).get('stability', '')})
        except InterruptedError as e:
            self.after(0, lambda e=e: messagebox.showinfo("Остановлено", str(e)))
        except Exception as e:
            logger.error(str(e))
            self.after(0, lambda e=e: messagebox.showerror("Ошибка", f"Ошибка генерации:\n{str(e)}"))
        finally:
            self._running = False; self.progress_var.set("Готов"); self._sp(0)
            if self.stop_btn and self.stop_btn.winfo_exists(): self.stop_btn.destroy()

    def update_scenario(self):
        if self.report_data: self.start_generation()

    def show_scenarios(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="Сценарии «что если»", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        fr = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8); fr.pack(fill=tk.X, padx=20, pady=10)
        ctk.CTkLabel(fr, text="Оптимистичный: спрос +").grid(row=0, column=0, padx=10, pady=6, sticky="w")
        ctk.CTkEntry(fr, textvariable=self.opt_pct, width=70).grid(row=0, column=1, padx=5)
        ctk.CTkLabel(fr, text="%   Пессимистичный: спрос −").grid(row=0, column=2, padx=10)
        ctk.CTkEntry(fr, textvariable=self.pes_pct, width=70).grid(row=0, column=3, padx=5)
        ctk.CTkLabel(fr, text="%").grid(row=0, column=4, padx=5)
        for s, l in [("optimistic", "Оптимистичный"), ("realistic", "Реалистичный"), ("pessimistic", "Пессимистичный")]:
            rf = ctk.CTkFrame(self.main_frame, fg_color="transparent"); rf.pack(fill=tk.X, padx=20, pady=3)
            ctk.CTkRadioButton(rf, text=l, variable=self.scenario, value=s).pack(side=tk.LEFT, padx=10)
        ctk.CTkButton(self.main_frame, text="Применить сценарий", command=self.start_generation, fg_color="#0057b3").pack(pady=20)

    def show_full_report(self):
        if not self.report_data: messagebox.showinfo("Инфо", "Сначала сгенерируйте отчёт."); return
        ReportViewer(self, self.report_data).focus()

    def show_recommendations(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="Рекомендации (пересчёт плана в реальном времени)",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        if not self.report_data:
            ctk.CTkLabel(self.main_frame, text="Сначала сгенерируйте отчёт.").pack(pady=20); return
        acts = self.report_data['actions']
        self.rec_vars = [BooleanVar(value=bool(a['roi']['delta_profit_month'] > 0)) for a in acts]
        cf = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8)
        cf.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        for a, v in zip(acts, self.rec_vars):
            r = a['roi']
            ctk.CTkCheckBox(cf, text=f"{a['Тип']} — {a['Товар']} (эффект {r['delta_profit_month']:+,.0f} ₽/мес, "
                                     f"капитал {r['capital_required']:,.0f} ₽)",
                            variable=v, command=self._plan_panel).pack(anchor="w", padx=20, pady=2)
            ctk.CTkLabel(cf, text=a['Обоснование'], font=ctk.CTkFont(size=10), text_color="gray",
                         wraplength=900, justify="left").pack(anchor="w", padx=40, pady=(0, 4))
        self.pp = ctk.CTkFrame(self.main_frame, fg_color="#eaf3fb", corner_radius=8); self.pp.pack(fill=tk.X, padx=20, pady=10)
        self.pl_label = ctk.CTkLabel(self.pp, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color="#0057b3")
        self.pl_label.pack(pady=8)
        b = ctk.CTkFrame(self.main_frame, fg_color="transparent"); b.pack(pady=10)
        ctk.CTkButton(b, text="Выбрать все", command=self._all).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Сбросить", command=self._none).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(b, text="Отчёт по выбранным рекомендациям", command=self._open_plan, fg_color="#0057b3").pack(side=tk.LEFT, padx=5)
        self._plan_panel()
    def _all(self):
        for v in self.rec_vars: v.set(True)
        self._plan_panel()
    def _none(self):
        for v in self.rec_vars: v.set(False)
        self._plan_panel()
    def _plan_panel(self):
        if not self.report_data: return
        mask = [bool(v.get()) for v in self.rec_vars]; d = self.report_data
        fut = sum(p['Прогнозная чистая прибыль'] if m else p['Текущая чистая прибыль'] for p, m in zip(d['profit_forecast'], mask))
        cap = sum(a['roi']['capital_required'] for a, m in zip(d['actions'], mask) if m)
        eff = sum(a['roi']['delta_profit_month'] for a, m in zip(d['actions'], mask) if m)
        self.pl_label.configure(text=f"Выбрано {sum(mask)} из {len(mask)} | Прибыль: {d['total_now']:,.0f} ₽ → "
                                     f"{fut:,.0f} ₽ | Эффект {eff:+,.0f} ₽/мес | Капитал: {cap:,.0f} ₽")
    def _open_plan(self):
        mask = [bool(v.get()) for v in self.rec_vars]
        ReportViewer(self, build_plan_report(self.report_data, mask)).focus()

    def show_history(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="История проектов", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        pr = self.db.get_projects()
        if not pr: ctk.CTkLabel(self.main_frame, text="Нет сохранённых проектов.").pack(pady=20); return
        for p in pr:
            f = ctk.CTkFrame(self.main_frame, fg_color="#f8f9fa", corner_radius=8); f.pack(fill=tk.X, padx=10, pady=5)
            ctk.CTkLabel(f, text=f"{p[1]} ({p[2][:10]})", font=ctk.CTkFont(size=13)).pack(side=tk.LEFT, padx=10)
            if p[5]:
                try:
                    k = json.loads(p[5])
                    ctk.CTkLabel(f, text=f"{k.get('net_profit_now', 0):,.0f} ₽ → {k.get('net_profit_forecast', 0):,.0f} ₽ "
                                         f"({k.get('change_pct', 0):+.1f}%)", font=ctk.CTkFont(size=10),
                                 text_color="gray").pack(side=tk.LEFT, padx=10)
                except Exception: pass
            ctk.CTkButton(f, text="Открыть", command=lambda i=p[0]: self._load(i)).pack(side=tk.RIGHT, padx=5)
            ctk.CTkButton(f, text="Удалить", command=lambda i=i0 if False else p[0]: self._del(i)).pack(side=tk.RIGHT, padx=5)
    def _load(self, i):
        d = self.db.load_project(i)
        if d: self.report_data = d; self.show_dashboard(); messagebox.showinfo("Успех", "Проект загружен.")
    def _del(self, i):
        if messagebox.askyesno("Удаление", "Удалить проект?"): self.db.delete_project(i); self.show_history()

    def show_about(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="О программе", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        for l in [f"Версия: {CONFIG['version']}", "",
                  "Все входные параметры — из файла данных или из настроек пользователя.",
                  "Цены сырья не задаются по умолчанию: только файл или ручной ввод.",
                  "Эластичность — из данных (R² > 0.1) либо из настроек с пометкой источника.",
                  "Полная методика и допущения — во вкладке «Допущения и методика» отчёта."]:
            ctk.CTkLabel(self.main_frame, text=l, font=ctk.CTkFont(size=12)).pack(anchor="w", padx=20)

    def show_help(self):
        self.clear_main()
        ctk.CTkLabel(self.main_frame, text="Помощь", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)
        ctk.CTkLabel(self.main_frame, text=f"Email: {CONFIG['support_email']}", font=ctk.CTkFont(size=12),
                     text_color="#0057b3").pack()

    def _onboarding(self):
        w = ctk.CTkToplevel(self); w.title("Первоначальная настройка"); w.geometry("620x520")
        w.transient(self); w.grab_set()
        ctk.CTkLabel(w, text="Первоначальная настройка", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        s = ctk.CTkFrame(w, fg_color="#f8f9fa", corner_radius=8); s.pack(fill=tk.X, padx=20, pady=10)
        entries = {}
        for lbl, dflt in [("Целевая чистая маржа, %", "20"), ("План роста выпуска, %", "20"),
                          ("Постоянные расходы, ₽/мес", "500000")]:
            ctk.CTkLabel(s, text=lbl).pack(anchor="w", padx=10)
            e = ctk.CTkEntry(s, width=140); e.insert(0, dflt); e.pack(anchor="w", padx=10, pady=2)
            entries[lbl] = e
        def fin():
            try:
                self.target_margin.set(entries["Целевая чистая маржа, %"].get().strip() or "20")
                self.increase_pct.set(entries["План роста выпуска, %"].get().strip() or "20")
                self.fixed_costs.set(float(entries["Постоянные расходы, ₽/мес"].get().strip() or 500000))
            except Exception:
                pass
            if not self.file_path.get(): self.file_path.set(generate_demo_data())
            w.destroy(); open("onboarding_done.txt", "w").close(); self.start_generation()
        ctk.CTkButton(w, text="Сгенерировать отчёт", command=fin, fg_color="#0057b3",
                      font=ctk.CTkFont(size=14, weight="bold")).pack(pady=20)

def run_selftest():
    for cols in [["Дата", "Товар", "Компоненты", "Количество единиц", "Цена за штуку"],
                 ["date", "product", "components", "quantity", "price"],
                 ["Дата продажи", "Наименование", "Состав", "Кол-во", "Цена, руб"]]:
        df = pd.DataFrame(columns=cols)
        rn = df.rename(columns={o: t for t, o in detect_columns(df).items()})
        for rq in ["Дата", "Товар", "Количество единиц", "Цена за штуку"]:
            assert rq in rn.columns, f"FAIL detector: {rq}"
    print("OK: детектор колонок")
    dbp = os.path.join(tempfile.gettempdir(), "fs_selftest.db")
    try:
        if os.path.exists(dbp): os.remove(dbp)
    except PermissionError: dbp = dbp.replace(".db", "_2.db")
    old = sqlite3.connect(dbp)
    old.execute("CREATE TABLE projects (id INTEGER PRIMARY KEY, name TEXT, date TEXT, data BLOB, notes TEXT)")
    old.commit(); old.close()
    db = HistoryDB(dbp); db.save_project("t", {"a": 1}, kpi_snapshot={"x": 1})
    assert len(db.get_projects()) == 1
    db.close()
    for _ in range(5):
        try: os.remove(dbp); break
        except PermissionError: import time as _t; _t.sleep(0.2)
        except FileNotFoundError: break
    print("OK: миграция БД")
    df_noprices = pd.DataFrame({"Дата": pd.date_range("2025-01-01", periods=40),
                                "Товар": ["A"] * 40, "Компоненты": ["X 1"] * 40,
                                "Количество единиц": [5] * 40, "Цена за штуку": [100.0] * 40})
    try:
        generate_report_data(df_noprices, params={})
        raise AssertionError("FAIL: обнаружены цены по умолчанию")
    except ValueError as e:
        assert "Цены сырья" in str(e)
    print("OK: запрет цен по умолчанию")
    rep = generate_report_data(generate_demo_data(), params={})
    acc = rep['best_model']['accuracy']
    if acc is not None:
        assert acc < 99, f"FAIL: подозрение на утечку цели ({acc}%)"
    assert all(a['Новые_шт'] >= 0 and a['Новая_цена'] > 0 for a in rep['actions'])
    for p, e in rep['product_economics'].items():
        assert e['break_even']['is_profitable'] == (e['pl']['net_profit'] > 0), f"FAIL BE: {p}"
    n = len(rep['actions'])
    assert abs(build_plan_report(rep, [True]*n)['total_future'] - rep['total_future']) < 1
    assert abs(build_plan_report(rep, [False]*n)['total_future'] - rep['total_now']) < 1
    assert rep['assumptions'] and rep['methodology'] and len(rep['scenario_table']) == 3
    for f in rep['forecasts'].values():
        assert f['forecast_std'] >= 0
        if f['method'] == 'среднее историческое':
            assert f['accuracy'] is None
    print(f"OK: прогон (точность {acc if acc is not None else 'н/д'}, изменение {rep['change_pct']:+.1f}%)")
    print("OK: все самопроверки пройдены")

if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        run_selftest()
    else:
        app = FinScopeApp()
        app.mainloop()
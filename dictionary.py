import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
import difflib
import sys
import os
import webbrowser

# Helper for PyInstaller / dev paths
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Load dictionary
def load_dictionary(file_path='dictionary.xlsx'):
    file_path = resource_path(file_path)
    try:
        df = pd.read_excel(file_path)
        df.columns = ['English', 'Eald-vacha', 'Notes'] if len(df.columns) == 3 else df.columns
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load dictionary: {e}")
        return None

# Pre-build all possible root terms
def build_roots(df):
    roots = set()
    for entry in df['Eald-vacha']:
        alts = [a.strip().lower() for a in str(entry).split('/')]
        for alt in alts:
            parts = [p.strip().lower() for p in alt.split('-') if p.strip()]
            roots.update(parts)
            roots.add(alt)
    return roots

# Get meaning for a term
def get_meaning(term, df):
    term_lower = term.lower()
    match = df[df['Eald-vacha'].str.lower() == term_lower]
    if not match.empty:
        return match['English'].values[0]
    for _, row in df.iterrows():
        alts = [a.strip().lower() for a in str(row['Eald-vacha']).split('/')]
        for alt in alts:
            parts = [p.strip().lower() for p in alt.split('-') if p.strip()]
            if term_lower in parts or term_lower == alt:
                return row['English']
    return "[unknown]"

# Find all segmentations
def find_segmentations(word_lower, roots, start=0, path=None):
    if path is None:
        path = []
    if start == len(word_lower):
        return [path[:]] if len(path) >= 2 else []
    segmentations = []
    for end in range(start + 3, len(word_lower) + 1):
        prefix = word_lower[start:end]
        if prefix in roots:
            path.append((start, end, word_lower[start:end]))
            segmentations.extend(find_segmentations(word_lower, roots, end, path))
            path.pop()
    return segmentations

# Score a segmentation
def score_segmentation(seg, word_lower, df, actual_eng):
    num_parts = len(seg)
    coverage = sum(end - start for start, end, _ in seg) / len(word_lower)
    if coverage < 0.9:
        return 0
    composed = " ".join(get_meaning(p, df) for _, _, p in seg).lower()
    sim = difflib.SequenceMatcher(None, composed, actual_eng.lower()).ratio()
    score = (sim * 80) + (num_parts * 10) + (coverage * 10)
    return score

# Possible decompositions
def find_possible_decompositions(word, df, roots):
    word_lower = word.lower()
    match = df[df['Eald-vacha'].str.lower() == word_lower]
    actual_eng = match['English'].values[0] if not match.empty else ""
    segmentations = find_segmentations(word_lower, roots)
    if not segmentations:
        return "No possible decompositions found."
    scored = []
    for seg in segmentations:
        score = score_segmentation(seg, word_lower, df, actual_eng)
        if score > 25:
            parts_str = " + ".join(f"{p}: {get_meaning(p, df)}" for _, _, p in seg)
            scored.append((score, parts_str))
    if not scored:
        return "No high-confidence decompositions."
    scored.sort(reverse=True)
    output = ""
    for i, (score, decomp) in enumerate(scored[:3], 1):
        output += f"{i}. (score: {int(score)}%) {decomp}\n"
    return output

# Decomposition
def decompose_word(word, df, indent=0, visited=None, roots=None):
    if visited is None:
        visited = set()
    if roots is None:
        roots = build_roots(df)
    word_lower = word.lower()
    if word_lower in visited:
        return f"{'  ' * indent}{word}: [cycle detected]"
    visited.add(word_lower)
    indent_str = '  ' * indent
    output = []
    if '-' in word:
        parts = [p.strip() for p in word.split('-')]
        decomp_parts = []
        for p in parts:
            sub = decompose_word(p, df, indent + 1, visited.copy(), roots)
            decomp_parts.append(sub)
        match = df[df['Eald-vacha'].str.lower() == word_lower]
        meaning = ""
        if not match.empty:
            eng = match['English'].values[0]
            notes = f" ({match['Notes'].values[0]})" if pd.notna(match['Notes'].values[0]) else ""
            meaning = f" → {eng}{notes}"
        output.append(f"{indent_str}{word} (compound){meaning}:\n" + '\n'.join(decomp_parts))
        return '\n'.join(output)
    if '/' in word:
        parts = [p.strip() for p in word.split('/')]
        decomp_parts = [decompose_word(p, df, indent, visited.copy(), roots) for p in parts]
        return '\n'.join(decomp_parts)
    if word_lower.startswith('nə'):
        base = word[2:].strip()
        if base:
            sub = decompose_word(base, df, indent + 1, visited.copy(), roots)
            output.append(f"{indent_str}{word} (negation prefix):\n{indent_str}  nə: not / negation / without\n{sub}")
            return '\n'.join(output)
    atomic_meaning = ""
    match = df[df['Eald-vacha'].str.lower() == word_lower]
    if not match.empty:
        eng = match['English'].values[0]
        notes = f" ({match['Notes'].values[0]})" if pd.notna(match['Notes'].values[0]) else ""
        atomic_meaning = f"{indent_str}{word}: {eng}{notes}"
    else:
        for _, row in df.iterrows():
            terms = [t.strip().lower() for t in str(row['Eald-vacha']).split('/')]
            if word_lower in terms:
                eng = row['English']
                notes = f" ({row['Notes']})" if pd.notna(row['Notes']) else ""
                atomic_meaning = f"{indent_str}{word}: {eng}{notes}"
                break
    if atomic_meaning:
        output.append(atomic_meaning)
    else:
        output.append(f"{indent_str}{word}: [not found]")
    if len(word) > 6 and '-' not in word and '/' not in word and not word_lower.startswith('nə'):
        possible = find_possible_decompositions(word, df, roots)
        if possible and "No" not in possible:
            output.append(f"{indent_str}Possible compound word roots:\n{possible}")
    return '\n'.join(output)

# Search functions
def search_word_exact_wildcard(df, query, direction):
    results = []
    original_query = query.strip()
    q = original_query.lower()
    if direction == 'English to Eald-vacha':
        search_col, result_col = 'English', 'Eald-vacha'
    else:
        search_col, result_col = 'Eald-vacha', 'English'
    starts_star = q.startswith('*')
    ends_star = q.endswith('*')
    if starts_star and ends_star:
        match_str = q[1:-1].strip()
        match_func = lambda t: match_str in t
    elif starts_star:
        match_str = q[1:].strip()
        match_func = lambda t: t.endswith(match_str)
    elif ends_star:
        match_str = q[:-1].strip()
        match_func = lambda t: t.startswith(match_str)
    else:
        match_str = q.strip()
        match_func = lambda t: t == match_str
    exclude_terms = ["habban", "amabba"]
    trigger = "abba"
    is_wildcard = starts_star or ends_star
    trigger_in_pattern = trigger in match_str
    pattern_has_excluded_term = any(excl in match_str for excl in exclude_terms)
    apply_exclusion = is_wildcard and trigger_in_pattern and not pattern_has_excluded_term
    for _, row in df.iterrows():
        terms = [t.strip().lower() for t in str(row[search_col]).split('/')]
        matched = False
        for term in terms:
            if apply_exclusion and any(excl in term for excl in exclude_terms):
                continue
            if match_func(term):
                matched = True
                break
        if matched:
            trans = row[result_col]
            notes = row['Notes'] if pd.notna(row['Notes']) else "No notes available."
            results.append(
                f"Match found for '{original_query}' in '{row[search_col]}':\n"
                f"{result_col}: {trans}\n"
                f"Notes: {notes}\n"
            )
    return results

def search_fuzzy(df, query, direction, min_score=75, limit=4):
    q_clean = query.strip().lower()
    if not q_clean:
        return []
    if direction == 'English to Eald-vacha':
        search_col, result_col = 'English', 'Eald-vacha'
    else:
        search_col, result_col = 'Eald-vacha', 'English'
    candidates = []
    for _, row in df.iterrows():
        terms = [t.strip() for t in str(row[search_col]).split('/')]
        for term in terms:
            score = difflib.SequenceMatcher(None, q_clean, term.lower()).ratio() * 100
            if score >= min_score:
                candidates.append((score, term, row[result_col], row['Notes'], row[search_col]))
    candidates.sort(key=lambda x: x[0], reverse=True)
    seen_rows = set()
    results = []
    for score, term, trans, notes_raw, orig_entry in candidates:
        row_id = id(orig_entry)
        if row_id in seen_rows:
            continue
        seen_rows.add(row_id)
        notes = notes_raw if pd.notna(notes_raw) else "No notes available."
        results.append(
            f"**Fuzzy match** (score: {int(score)}%): '{term}' in '{orig_entry}'\n"
            f"{result_col}: {trans}\n"
            f"Notes: {notes}\n"
        )
        if len(results) >= limit:
            break
    return results

# GUI with Help menu and keyboard shortcuts
class DictionaryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bilingual Dictionary: English ↔ Eald-vacha")
        self.root.geometry("750x700")

        self.df = load_dictionary()
        if self.df is None:
            self.root.quit()

        self.roots = build_roots(self.df)

        self.last_results = []
        self.last_direction = ''
        self.last_query = ''

        # Add Help menu at the top
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu, underline=0)
        help_menu.add_command(label="Show Help", command=self.show_help, underline=0)

        # Input frame
        input_frame = ttk.Frame(root)
        input_frame.pack(pady=10, padx=15, fill='x')
        ttk.Label(input_frame, text="Search (wildcards * supported):").pack(side='left', padx=(0, 10))
        self.entry = ttk.Entry(input_frame, width=55, font=('Consolas', 11))
        self.entry.pack(side='left', expand=True, fill='x')
        self.entry.focus_set()

        # Direction
        dir_frame = ttk.Frame(root)
        dir_frame.pack(pady=5)
        self.direction = tk.StringVar(value='English to Eald-vacha')
        ttk.Radiobutton(dir_frame, text="English → Eald-vacha", variable=self.direction, value='English to Eald-vacha').pack(side='left', padx=12)
        ttk.Radiobutton(dir_frame, text="Eald-vacha → English", variable=self.direction, value='Eald-vacha to English').pack(side='left', padx=12)

        # Fuzzy tolerance slider
        fuzzy_frame = ttk.Frame(root)
        fuzzy_frame.pack(pady=8, padx=15, fill='x')
        ttk.Label(fuzzy_frame, text="Fuzzy match tolerance:").pack(side='left', padx=(0, 10))
        self.fuzzy_var = tk.DoubleVar(value=75.0)
        slider = ttk.Scale(fuzzy_frame, from_=50, to=95, orient='horizontal', variable=self.fuzzy_var, length=250)
        slider.pack(side='left', padx=5)
        self.tolerance_label = ttk.Label(fuzzy_frame, text="75%")
        self.tolerance_label.pack(side='left', padx=(10, 0))

        def update_label(*args):
            self.tolerance_label.config(text=f"{int(self.fuzzy_var.get())}%")
        self.fuzzy_var.trace('w', update_label)

        # Buttons (added new "Add nə" button)
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text="Search (Alt+S)", command=self.perform_search, width=12).pack(side='left', padx=5)
        self.copy_btn = ttk.Button(btn_frame, text="Copy (Alt+C)", command=self.copy_to_clipboard, state='disabled')
        self.copy_btn.pack(side='left', padx=5)
        self.decomp_btn = ttk.Button(btn_frame, text="Decompose (Alt+D)", command=self.perform_decompose, state='disabled')
        self.decomp_btn.pack(side='left', padx=5)
        # NEW button: Add nə
        ttk.Button(btn_frame, text="Add nə (Alt+P)", command=self.insert_nə).pack(side='left', padx=5)

        # Output
        self.output = tk.Text(root, height=22, width=85, wrap='word', font=('Consolas', 10))
        self.output.pack(pady=10, padx=15, fill='both', expand=True)

        scrollbar = ttk.Scrollbar(root, orient='vertical', command=self.output.yview)
        scrollbar.pack(side='right', fill='y')
        self.output.configure(yscrollcommand=scrollbar.set)

        self.has_results = False

        # ── Keyboard Shortcuts ──────────────────────────────────────────────
        self.root.bind('<Alt-s>', lambda e: self.perform_search())
        self.root.bind('<Alt-S>', lambda e: self.perform_search())
        self.root.bind('<Alt-d>', lambda e: self.perform_decompose_safe())
        self.root.bind('<Alt-D>', lambda e: self.perform_decompose_safe())
        self.root.bind('<Alt-n>', lambda e: self.set_direction('English to Eald-vacha'))
        self.root.bind('<Alt-N>', lambda e: self.set_direction('English to Eald-vacha'))
        self.root.bind('<Alt-v>', lambda e: self.set_direction('Eald-vacha to English'))
        self.root.bind('<Alt-V>', lambda e: self.set_direction('Eald-vacha to English'))
        self.root.bind('<Alt-e>', lambda e: self.focus_search())
        self.root.bind('<Alt-E>', lambda e: self.focus_search())
        self.root.bind('<Alt-c>', lambda e: self.copy_to_clipboard_safe())
        self.root.bind('<Alt-C>', lambda e: self.copy_to_clipboard_safe())
        self.root.bind('<Alt-x>', lambda e: self.close_program())
        self.root.bind('<Alt-X>', lambda e: self.close_program())
        self.root.bind('<Alt-plus>', lambda e: self.adjust_fuzzy(5))
        self.root.bind('<Alt-minus>', lambda e: self.adjust_fuzzy(-5))
        self.entry.bind('<Return>', lambda e: self.perform_search())

        # NEW: Alt + P → Insert "nə" into search box
        self.root.bind('<Alt-p>', lambda e: self.insert_nə())
        self.root.bind('<Alt-P>', lambda e: self.insert_nə())

        # Alt + H → show help
        self.root.bind('<Alt-h>', lambda e: self.show_help())
        self.root.bind('<Alt-H>', lambda e: self.show_help())

    def insert_nə(self):
        """Insert 'nə' at the current cursor position in the search box (Alt + P)"""
        current_text = self.entry.get()
        cursor_pos = self.entry.index(tk.INSERT)
        new_text = current_text[:cursor_pos] + "nə" + current_text[cursor_pos:]
        self.entry.delete(0, tk.END)
        self.entry.insert(0, new_text)
        # Move cursor after the inserted "nə"
        self.entry.icursor(cursor_pos + 2)
        self.entry.focus_set()

    def adjust_fuzzy(self, delta):
        current = self.fuzzy_var.get()
        new_value = max(50, min(95, current + delta))
        self.fuzzy_var.set(new_value)
        self.tolerance_label.config(text=f"{int(new_value)}%")

    def show_help(self):
        help_text = (
            "Eald-vacha Bilingual Dictionary 1.0 (Released: 4 January 2026)\n\n"
            "Purpose:\n"
            "This program is a bidirectional dictionary for English and the constructed language Eald-vacha. "
            "It supports exact matches, wildcard searches (*), fuzzy/typo-tolerant search, and advanced morphological decomposition of compound words.\n\n"
            "Keyboard Shortcuts:\n"
            "- Alt + S: Perform Search\n"
            "- Alt + D: Decompose Results (English → Eald-vacha only, after a successful search)\n"
            "- Alt + N: Switch to English → Eald-vacha\n"
            "- Alt + V: Switch to Eald-vacha → English\n"
            "- Alt + E: Focus/return to search bar\n"
            "- Alt + C: Copy Results to Clipboard\n"
            "- Alt + X: Close the program (with confirmation)\n"
            "- Alt + +: Increase fuzzy tolerance by 5%\n"
            "- Alt + -: Decrease fuzzy tolerance by 5%\n"
            "- Alt + P: Insert 'nə' prefix into search box\n"
            "- Enter (in search box): Perform Search\n"
            "- Alt + H: Show this Help window\n\n"
            "Please note that the Decompose Results feature only works for English → Eald-vacha searches.\n\n"
            "Developed by: Brant von Goble (Eald-vacha-abba)\n\n"
            "Licensing:\n"
            "This program is licensed under the Creative Commons Attribution 4.0 International (CC BY 4.0) license.\n"
            "You are free to share and adapt it, even commercially, as long as you give appropriate credit.\n"
            "Full license: https://creativecommons.org/licenses/by/4.0/\n\n"
            "Enjoy exploring Eald-vacha!"
        )

        help_window = tk.Toplevel(self.root)
        help_window.title("Help - Eald-vacha Dictionary")
        help_window.geometry("600x550")
        help_window.transient(self.root)
        help_window.grab_set()

        text_widget = tk.Text(help_window, wrap='word', font=('Consolas', 10))
        text_widget.pack(padx=15, pady=15, fill='both', expand=True)

        scrollbar = ttk.Scrollbar(help_window, orient='vertical', command=text_widget.yview)
        scrollbar.pack(side='right', fill='y')
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.insert(tk.END, help_text)
        text_widget.configure(state='disabled')

        license_label = ttk.Label(help_window, text="Click here to open full CC BY 4.0 license", foreground="blue", cursor="hand2")
        license_label.pack(pady=10)
        license_label.bind("<Button-1>", lambda e: webbrowser.open("https://creativecommons.org/licenses/by/4.0/"))

        ttk.Button(help_window, text="Close (Alt+B)", command=help_window.destroy).pack(pady=10)

        # Alt + B to close Help window
        help_window.bind('<Alt-b>', lambda e: help_window.destroy())
        help_window.bind('<Alt-B>', lambda e: help_window.destroy())

        help_window.focus_force()

    def set_direction(self, direction_value):
        self.direction.set(direction_value)
        self.entry.focus_set()

    def perform_decompose_safe(self):
        if not self.last_results or self.last_direction != 'English to Eald-vacha':
            messagebox.showinfo("Not Available",
                               "Decomposition is only available after a successful English → Eald-vacha search.")
            return
        self.perform_decompose()

    def perform_search(self):
        query = self.entry.get().strip()
        if not query:
            messagebox.showwarning("Input Required", "Please enter something to search.")
            return
        self.last_query = query
        self.output.delete(1.0, tk.END)
        dir_val = self.direction.get()
        self.last_direction = dir_val
        min_score = int(self.fuzzy_var.get())
        results = search_word_exact_wildcard(self.df, query, dir_val)
        self.last_results = results
        if not results:
            fuzzy_results = search_fuzzy(self.df, query, dir_val, min_score=min_score)
            self.last_results = fuzzy_results
            if fuzzy_results:
                self.output.insert(tk.END, f"No exact/wildcard match for '{query}'.\n\n"
                                          f"Fuzzy matches (min similarity {min_score}%):\n\n")
                self.output.insert(tk.END, "\n---\n".join(fuzzy_results) + "\n")
            else:
                self.output.insert(tk.END, f"No matches found (exact, wildcard, or fuzzy ≥ {min_score}%).\n")
        else:
            self.output.insert(tk.END, "\n---\n".join(results) + "\n")
        self.has_results = bool(self.output.get(1.0, tk.END).strip())
        self.copy_btn.config(state='normal' if self.has_results else 'disabled')
        self.decomp_btn.config(state='normal' if self.has_results and dir_val == 'English to Eald-vacha' else 'disabled')
        self.output.see(tk.END)

    def perform_decompose(self):
        self.output.insert(tk.END, f"\n=== Decompositions ===\n\n")
        for result in self.last_results:
            lines = result.split('\n')
            word_line = next((l for l in lines if l.startswith('Eald-vacha: ')), None)
            if word_line:
                word = word_line[len('Eald-vacha: '):].strip()
                decomp = decompose_word(word, self.df, roots=self.roots)
                self.output.insert(tk.END, f"Decomposition for {word}:\n{decomp}\n\n")
        self.output.see(tk.END)

    def copy_to_clipboard(self):
        if self.has_results:
            text = self.output.get(1.0, tk.END).strip()
            try:
                pyperclip.copy(text)
                messagebox.showinfo("Success", "Results copied to clipboard!")
            except Exception as e:
                messagebox.showerror("Clipboard Error", f"Failed to copy: {e}")

    def focus_search(self):
        self.entry.focus_set()

    def copy_to_clipboard_safe(self):
        if self.has_results:
            self.copy_to_clipboard()
        else:
            messagebox.showinfo("Nothing to Copy", "No results available to copy.")

    def close_program(self):
        if messagebox.askyesno("Exit", "Are you sure you want to close the program?"):
            self.root.quit()

    def adjust_fuzzy(self, delta):
        current = self.fuzzy_var.get()
        new_value = max(50, min(95, current + delta))
        self.fuzzy_var.set(new_value)
        self.tolerance_label.config(text=f"{int(new_value)}%")

if __name__ == "__main__":
    root = tk.Tk()
    app = DictionaryApp(root)
    root.mainloop()
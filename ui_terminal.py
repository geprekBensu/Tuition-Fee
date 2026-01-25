import os
import sys
import pandas as pd
import PySimpleGUI as sg
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


def read_data(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xls', '.xlsx'):
        return pd.read_excel(path)
    if ext in ('.csv', '.txt'):
        return pd.read_csv(path)
    raise ValueError('Unsupported file type: ' + ext)


def draw_figure(canvas, figure):
    if canvas is None:
        return None
    figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
    figure_canvas_agg.draw()
    figure_canvas_agg.get_tk_widget().pack(side='top', fill='both', expand=1)
    return figure_canvas_agg


def make_layout():
    left_col = [
        [sg.Text('Load Excel/CSV file:'), sg.Input(key='-FILE-'), sg.FileBrowse(file_types=(('Excel', '*.xls;*.xlsx'), ('CSV', '*.csv')))],
        [sg.Button('Load', key='-LOAD-'), sg.Button('Exit')],
        [sg.Frame('Preview (first 100 rows)', [[sg.Table(values=[], headings=[], auto_size_columns=True, display_row_numbers=False, num_rows=15, key='-TABLE-')]])],
    ]

    right_col = [
        [sg.Frame('Summary', [[sg.Multiline('', size=(40,10), key='-SUMMARY-', disabled=True)]])],
        [sg.Text('Plot Column:'), sg.Combo(values=[], key='-COLS-', size=(30,1)), sg.Button('Plot', key='-PLOT-')],
        [sg.Canvas(key='-CANVAS-')],
    ]

    layout = [[sg.Column(left_col), sg.VerticalSeparator(), sg.Column(right_col)]]
    return layout


def update_table(window, df):
    if df is None or df.empty:
        window['-TABLE-'].update(values=[], headings=[])
        return
    sample = df.head(100)
    headings = list(sample.columns.astype(str))
    values = sample.fillna('').values.tolist()
    window['-TABLE-'].update(values=values, headings=headings)


def update_summary(window, df):
    if df is None:
        window['-SUMMARY-'].update('')
        return
    desc = df.describe(include='all').transpose()
    txt = desc.to_string()
    window['-SUMMARY-'].update(txt)


def main():
    sg.theme('DefaultNoMoreNagging')
    layout = make_layout()
    window = sg.Window('Data Terminal — Friendly UI', layout, finalize=True, resizable=True)

    figure_agg = None
    df = None

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        if event == '-LOAD-':
            path = values['-FILE-']
            if not path or not os.path.exists(path):
                sg.popup('Please choose a valid file to load.')
                continue
            try:
                df = read_data(path)
            except Exception as e:
                sg.popup('Failed to read file:', str(e))
                df = None
                continue
            update_table(window, df)
            update_summary(window, df)
            cols = list(df.select_dtypes(include=['number']).columns.astype(str))
            # include non-numeric columns too for quick plotting by index
            all_cols = list(df.columns.astype(str))
            window['-COLS-'].update(values=all_cols, value=(all_cols[0] if all_cols else ''))
        if event == '-PLOT-':
            if df is None:
                sg.popup('Load data first.')
                continue
            col = values['-COLS-']
            if not col or col not in df.columns:
                sg.popup('Choose a valid column to plot.')
                continue
            # clear previous figure
            if figure_agg:
                try:
                    figure_agg.get_tk_widget().forget()
                except Exception:
                    pass
            fig, ax = plt.subplots(figsize=(5, 3))
            try:
                series = pd.to_numeric(df[col], errors='coerce')
                series.plot(ax=ax)
                ax.set_title(f'{col} (index vs value)')
                ax.set_xlabel('index')
                ax.set_ylabel(col)
            except Exception:
                # fallback: bar plot of value counts
                df[col].value_counts().head(20).plot(kind='bar', ax=ax)
                ax.set_title(f'{col} (value counts)')
            figure_agg = draw_figure(window['-CANVAS-'].TKCanvas, fig)

    window.close()


if __name__ == '__main__':
    main()

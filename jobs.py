import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


def select_excel_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path

#chage for git
'''jhhsdkjhjsdjbb'''
def read_excel_file(file_path):
    if file_path:
        df = pd.read_excel(file_path)
        return df
    else:
        return None


def filter_companies(df):
    company_counts = df['Company'].value_counts().sort_index()
    company_counts_filtered = company_counts[company_counts >= 2]
    return company_counts_filtered


def calculate_rejection_counts(df, company_counts_filtered):
    rejection_counts = pd.Series(0, index=company_counts_filtered.index)
    rejections = df[df['Result'] == 'Rejection']['Company'].value_counts().sort_index()
    rejection_counts.update(rejections)
    return rejection_counts


def plot_data(company_counts_filtered, rejection_counts):
    fig, ax1 = plt.subplots(figsize=(12, 8))

    bars1 = ax1.bar(company_counts_filtered.index, company_counts_filtered.values, color='skyblue',
                    label='Number of Occurrences')
    ax1.set_ylabel('Number of Occurrences', color='skyblue')
    ax1.tick_params(axis='y', labelcolor='skyblue')
    plt.xticks(rotation=90, ha='right')

    ax2 = ax1.twinx()
    bars2 = ax2.bar(rejection_counts.index, rejection_counts.values, color='orange', label='Rejection Number')
    ax2.set_ylabel('Rejection Number', color='orange')
    ax2.tick_params(axis='y', labelcolor='orange')

    max_value = max(max(company_counts_filtered.values), max(rejection_counts.values))
    ax1.set_ylim(0, max_value)
    ax2.set_ylim(0, max_value)

    rejection_percentage = (rejection_counts / company_counts_filtered) * 100

    for i, (rect1, rect2) in enumerate(zip(bars1, bars2)):
        height1 = rect1.get_height()
        height2 = rect2.get_height()
        avg_height = (height1 + height2) / 2
        ax1.annotate(f'{rejection_percentage.iloc[i]:.1f}%',
                     xy=(rect1.get_x() + rect1.get_width() / 2, avg_height),
                     xytext=(0, 3),
                     textcoords="offset points",
                     ha='center', va='bottom',
                     rotation=90)

    plt.title('Number of Occurrences and Rejection Numbers by Company')
    fig.tight_layout()
    plt.show()


def plot_data_decending(company_counts_filtered, rejection_counts):
    fig, ax1 = plt.subplots(figsize=(12, 8))
    rejection_percentage = (rejection_counts / company_counts_filtered) * 100
    sorted_rejection_percentage = rejection_percentage.sort_values(ascending=False)

    fig, ax1 = plt.subplots(figsize=(12, 8))

    # Reorder the bars based on sorted rejection percentage
    company_counts_filtered_sorted = company_counts_filtered.loc[sorted_rejection_percentage.index]
    rejection_counts_sorted = rejection_counts.loc[sorted_rejection_percentage.index]

    bars1 = ax1.bar(company_counts_filtered_sorted.index, company_counts_filtered_sorted.values, color='skyblue',
                    label='Number of Occurrences')
    ax1.set_ylabel('Number of Occurrences', color='skyblue')
    ax1.tick_params(axis='y', labelcolor='skyblue')
    plt.xticks(rotation=90, ha='right')

    ax2 = ax1.twinx()
    bars2 = ax2.bar(rejection_counts_sorted.index, rejection_counts_sorted.values, color='orange',
                    label='Rejection Number')
    ax2.set_ylabel('Rejection Number', color='orange')
    ax2.tick_params(axis='y', labelcolor='orange')

    max_value = max(max(company_counts_filtered_sorted.values), max(rejection_counts_sorted.values))
    ax1.set_ylim(0, max_value)
    ax2.set_ylim(0, max_value)

    for i, (rect1, rect2) in enumerate(zip(bars1, bars2)):
        height1 = rect1.get_height()
        height2 = rect2.get_height()
        avg_height = (height1 + height2) / 2
        ax1.annotate(f'{sorted_rejection_percentage.iloc[i]:.1f}%',
                     xy=(rect1.get_x() + rect1.get_width() / 2, avg_height),
                     xytext=(0, 3),
                     textcoords="offset points",
                     ha='center', va='bottom',
                     rotation=90)

    plt.title('Number of Occurrences and Rejection Numbers by Company (Descending Order of Rejection Percentage)')
    fig.tight_layout()
    plt.show()


def plot_data_number(company_counts_filtered):
    # Sort the company counts in descending order
    company_counts_sorted = company_counts_filtered.sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(12, 8))

    # Plotting the number of occurrences for each company
    bars = ax.bar(company_counts_sorted.index, company_counts_sorted.values, color='skyblue')

    ax.set_ylabel('Number of Occurrences', color='skyblue')
    ax.tick_params(axis='y', labelcolor='skyblue')

    plt.xticks(rotation=90, ha='right')

    # Adding labels on top of each bar
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom',
                    rotation=90)

    plt.title('Number of Occurrences by Company (Descending Order)')
    fig.tight_layout()
    plt.show()


def main():
    # file_path = select_excel_file()
    file_path = "D:\Applications\A Application DOCS\Application tracker.xlsx"
    if file_path:
        df = read_excel_file(file_path)
        if df is not None:
            company_counts_filtered = filter_companies(df)
            rejection_counts = calculate_rejection_counts(df, company_counts_filtered)
            plot_data(company_counts_filtered, rejection_counts)
            plot_data_number(company_counts_filtered)
            plot_data_decending(company_counts_filtered, rejection_counts)
    else:
        print("No file selected.")


if __name__ == "__main__":
    main()

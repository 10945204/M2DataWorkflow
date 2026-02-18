import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os

def main():
    # Load data
    file_path = 'Grad Program Exit Survey Data 2024.xlsx'
    print(f"Loading data from {file_path}...")
    try:
        xl = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"Error: File {file_path} not found.")
        return

    # Parse with header at row index 1 (the question text)
    # Row 0 is the short code (Q35_1 etc), Row 1 is the question text.
    df = xl.parse(xl.sheet_names[0], header=1)

    # Columns L-S correspond to indices 11-18.
    target_cols_indices = list(range(11, 19))
    target_cols = df.columns[target_cols_indices]

    # Extract course names
    # Header format: "... - Course Name"
    course_names = []
    for col in target_cols:
        if " - " in col:
            course_names.append(col.split(" - ")[-1])
        else:
            course_names.append(col)

    print("Identified courses:", course_names)

    # Select data
    # The dataframe 'df' has headers from Row 1.
    # The first row of data in 'df' (index 0) corresponds to Row 2 of the Excel file, which is the metadata row ({"ImportId":...}).
    # We need to skip this row.
    data = df.iloc[1:, target_cols_indices].copy()

    # Rename columns to course names
    data.columns = course_names

    # Convert to numeric and drop NaN
    # Errors='coerce' turns non-numeric (like the metadata if we hadn't skipped it, or other junk) into NaN.
    data = data.apply(pd.to_numeric, errors='coerce')

    initial_count = len(data)
    data.dropna(inplace=True)
    final_count = len(data)
    print(f"Data cleaning: {initial_count} rows -> {final_count} rows (dropped {initial_count - final_count} rows with missing values).")

    # Calculate stats
    stats = pd.DataFrame({
        'Average Rating': data.mean(),
        'Count': data.count()
    })

    # Sort
    # Rank 1 (best) has value 1. So lowest average is best.
    # Tie break: higher volume ranks higher (better).
    # So we want to sort by Average Rating (Ascending) and then Count (Descending).
    stats.sort_values(by=['Average Rating', 'Count'], ascending=[True, False], inplace=True)

    print("Ranking calculated:")
    print(stats)

    # Save CSV
    stats.reset_index(inplace=True)
    stats.rename(columns={'index': 'Course'}, inplace=True)

    output_dir = 'results'
    os.makedirs(output_dir, exist_ok=True)

    csv_path = os.path.join(output_dir, 'program_ranking.csv')
    stats.to_csv(csv_path, index=False)
    print(f"Saved ranking to {csv_path}")

    # Plot
    plt.figure(figsize=(12, 8))

    # We want the best course (top of list) to be at the top of the chart.
    # stats is sorted Best -> Worst.
    # We reverse it for plotting so Best is at top (y-axis max).
    plot_data = stats.iloc[::-1]

    bars = plt.barh(plot_data['Course'], plot_data['Average Rating'], color='#185C33')
    plt.xlabel('Average Rating (1 = Best, 8 = Worst)')
    plt.ylabel('MAcc CORE Courses')
    plt.title('Course Ranking by Student Preference')
    plt.grid(axis='x', linestyle='--', alpha=0.7)

    # Add value labels
    for bar in bars:
        width = bar.get_width()
        plt.text(width + 0.05, bar.get_y() + bar.get_height()/2, f'{width:.2f}',
                 va='center', fontsize=10)

    plt.tight_layout()
    png_path = os.path.join(output_dir, 'program_ranking.png')
    plt.savefig(png_path)
    print(f"Saved plot to {png_path}")

if __name__ == "__main__":
    main()

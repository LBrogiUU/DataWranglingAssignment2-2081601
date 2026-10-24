
import pandas as pd
from openpyxl.utils import get_column_letter


# Exact column order from the original COPOD paper
ALGORITHMS = [
    "ABOD", "CBLOF", "FB", "HBOS", "IForest",
    "KNN", "LODA", "LOF", "OCSVM", "COPOD"
]

COMPARISONS = {
    "ROC_AUC": (
        "COPOD_paper_ROC_AUC.csv",
        "ROC_AUC_merged.csv"
    ),
    "AP": (
        "COPOD_paper_AP.csv",
        "AP_merged.csv"
    )
}


def load_scores(filename):
    df = pd.read_csv(filename, dtype=str)
    df.columns = df.columns.str.strip()

    # Keep original dataset names for display
    df["Data"] = df["Data"].str.strip()

    # Remove rankings and convert scores to numbers
    for algorithm in ALGORITHMS:
        df[algorithm] = (
            df[algorithm]
            .str.extract(r"([-+]?\d*\.?\d+)")[0]
            .astype(float)
        )

    # Standardize dataset names for matching
    df["match_name"] = df["Data"].str.lower()

    return df.set_index("match_name")


def format_difference(value):
    # Round to 4 decimal places
    value = round(float(value), 4)

    # Display zero as "0"
    if value == 0:
        return "0"

    # Remove unnecessary trailing zeros
    return f"{value:.4f}".rstrip("0").rstrip(".")


def compare_results(metric, paper_file, reproduced_file):
    paper = load_scores(paper_file)
    reproduced = load_scores(reproduced_file)

    # Check that datasets match
    missing = paper.index.difference(reproduced.index)
    extra = reproduced.index.difference(paper.index)

    if len(missing) or len(extra):
        raise ValueError(
            f"{metric}:\n"
            f"Missing datasets: {list(missing)}\n"
            f"Extra datasets: {list(extra)}"
        )

    # Ensure identical dataset order
    reproduced = reproduced.loc[paper.index]

    # Calculate signed differences
    # Positive = reproduction higher than paper
    # Negative = reproduction lower than paper
    difference = reproduced[ALGORITHMS] - paper[ALGORITHMS]

    # Keep reproduction dataset names
    result = difference.copy()
    result.insert(0, "Data", reproduced["Data"].values)

    # Format differences without rankings
    for algorithm in ALGORITHMS:
        result[algorithm] = result[algorithm].apply(
            format_difference
        )

    # Ensure same column order as reproduction
    result = result[["Data"] + ALGORITHMS]

    # Save CSV
    output_name = f"{metric}_absolute_difference"
    result.to_csv(f"{output_name}.csv", index=False)

    # Save Excel without bold formatting
    with pd.ExcelWriter(
        f"{output_name}.xlsx",
        engine="openpyxl"
    ) as writer:

        result.to_excel(
            writer,
            index=False,
            sheet_name="Results"
        )

        ws = writer.sheets["Results"]

        # Adjust column widths
        for column in ws.columns:
            max_length = max(
                len(str(cell.value))
                if cell.value is not None else 0
                for cell in column
            )

            ws.column_dimensions[
                get_column_letter(column[0].column)
            ].width = min(max_length + 3, 30)

    # Summary for COPOD, excluding AVG row
    copod_diff = difference.drop(index="avg")["COPOD"]

    print(f"\n{metric} - COPOD comparison:")
    print(
        f"Mean absolute difference: "
        f"{copod_diff.abs().mean():.4f}"
    )
    print(
        f"Maximum absolute difference: "
        f"{copod_diff.abs().max():.4f}"
    )
    print(
        f"Datasets with higher reproduced scores: "
        f"{(copod_diff > 0).sum()}"
    )
    print(
        f"Datasets with lower reproduced scores: "
        f"{(copod_diff < 0).sum()}"
    )

    print(
        f"Saved: {output_name}.csv "
        f"and {output_name}.xlsx"
    )


def main():
    for metric, (paper_file, reproduced_file) in COMPARISONS.items():
        compare_results(
            metric,
            paper_file,
            reproduced_file
        )

    print("\nDone! Both comparisons completed successfully.")


if __name__ == "__main__":
    main()

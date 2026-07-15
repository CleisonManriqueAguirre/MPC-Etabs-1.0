from __future__ import annotations

from io import StringIO
from typing import Optional

import pandas as pd

from . import etabs_connection


class EtabsTableExtractor:
    """Encapsulate ETABS table lookup, extraction, and formatting."""

    TABLE_HINTS = (
        (("force", "forces", "moment", "moments", "frame"), "Element Joint Forces - Frame"),
        (("displacement", "displacements", "deflection", "deflections"), "Joint Displacements"),
        (("reaction",), "Base Reactions"),
        (("modal", "mode", "period"), "Modal Periods And Frequencies"),
        (("design", "steel", "beam"), "Steel Frame Design Summary - AISC 360-16"),
        (("design", "steel"), "Steel Frame Design Summary - AISC 360-16"),
        (("design", "concrete", "column"), "Concrete Column Design Summary - ACI 318-19"),
        (("design", "concrete"), "Concrete Beam Design Summary - ACI 318-19"),
        (("joint", "coordinate", "coordinates"), "Objects and Elements - Joints"),
        (("frame",), "Objects and Elements - Frames"),
        (("material",), "Material Properties - Basic Mechanical Properties"),
        (("section",), "Frame Section Property Definitions - General"),
        (("load", "pattern"), "Load Pattern Definitions"),
        (("load", "combination"), "Load Combination Definitions"),
        (("area",), "Objects and Elements - Areas"),
        (("story", "stories"), "Story Definitions"),
        (("grid",), "Grid Definitions - General"),
    )

    def __init__(self, connection=etabs_connection) -> None:
        self._connection = connection

    def is_connected(self) -> bool:
        return self._connection.model is not None

    def etabs_extract(self, table_name: str) -> Optional[pd.DataFrame]:
        try:
            ret = self._connection.model.DatabaseTables.GetTableForDisplayCSVString(
                table_name, [], "", 0, "", ","
            )
            if ret[-1] == 0:
                csv_string = ret[2]
                if csv_string and len(csv_string) > 20:
                    return pd.read_csv(StringIO(csv_string))
            return None
        except Exception as exc:
            print(f"Table extraction failed for '{table_name}': {exc}")
            return None

    def find_table_name(self, description: str) -> Optional[str]:
        desc_lower = description.lower()

        for keywords, table_name in self.TABLE_HINTS:
            if all(term in desc_lower for term in keywords):
                return table_name

        return None

    def get_etabs_table(
        self,
        table_description: str,
        export_format: str = "dataframe",
        output_file: str | None = None,
    ) -> str:
        if not self.is_connected():
            return "Not connected to ETABS. Use connect_etabs() first."

        try:
            table_name = self.find_table_name(table_description)
            if not table_name:
                return (
                    f"Could not find table for: '{table_description}'\n"
                    "Try: 'frame forces', 'joint displacements', 'steel design'"
                )

            df = self.etabs_extract(table_name)
            if df is None:
                return f"No data in table '{table_name}'. Check if analysis/design has been run."

            return self.format_table_output(df, export_format, output_file, table_description)

        except Exception as exc:
            return f"Error extracting table: {exc}"

    def extract_joints_api(self) -> Optional[pd.DataFrame]:
        try:
            ret = self._connection.model.PointObj.GetNameList()
            if ret[0] != 0:
                return None

            joint_data = []
            for joint_name in ret[1]:
                try:
                    coord_ret = self._connection.model.PointObj.GetCoordCartesian(joint_name)
                    if coord_ret[0] != 0:
                        continue

                    restraint_ret = self._connection.model.PointObj.GetRestraint(joint_name)
                    restraints = restraint_ret[1] if restraint_ret[0] == 0 else [False] * 6

                    joint_data.append(
                        {
                            "Joint": joint_name,
                            "X": coord_ret[1],
                            "Y": coord_ret[2],
                            "Z": coord_ret[3],
                            "UX_Restraint": restraints[0],
                            "UY_Restraint": restraints[1],
                            "UZ_Restraint": restraints[2],
                            "RX_Restraint": restraints[3],
                            "RY_Restraint": restraints[4],
                            "RZ_Restraint": restraints[5],
                        }
                    )
                except Exception:
                    continue

            return pd.DataFrame(joint_data) if joint_data else None

        except Exception as exc:
            print(f"API joints extraction failed: {exc}")
            return None

    def extract_frames_api(self) -> Optional[pd.DataFrame]:
        try:
            ret = self._connection.model.FrameObj.GetNameList()
            if ret[0] != 0:
                return None

            frame_data = []
            for frame_name in ret[1]:
                try:
                    points_ret = self._connection.model.FrameObj.GetPoints(frame_name)
                    if points_ret[0] != 0:
                        continue

                    section_ret = self._connection.model.FrameObj.GetSection(frame_name)
                    material_ret = self._connection.model.FrameObj.GetMaterial(frame_name)
                    length_ret = self._connection.model.FrameObj.GetLength(frame_name)

                    frame_data.append(
                        {
                            "Frame": frame_name,
                            "Point1": points_ret[1],
                            "Point2": points_ret[2],
                            "Length": length_ret[1] if length_ret[0] == 0 else 0,
                            "Section": section_ret[1] if section_ret[0] == 0 else "Unknown",
                            "Material": material_ret[1] if material_ret[0] == 0 else "Unknown",
                        }
                    )
                except Exception:
                    continue

            return pd.DataFrame(frame_data) if frame_data else None

        except Exception as exc:
            print(f"API frames extraction failed: {exc}")
            return None

    def get_etabs_data_api(self, data_type: str) -> str:
        if not self.is_connected():
            return "Not connected to ETABS. Use connect_etabs() first."

        try:
            normalized = data_type.lower()
            if normalized == "joints":
                df = self.extract_joints_api()
                return f"API Joints Data:\n{self.create_dataframe_preview(df)}" if df is not None else "No joint data available via API"

            if normalized == "frames":
                df = self.extract_frames_api()
                return f"API Frames Data:\n{self.create_dataframe_preview(df)}" if df is not None else "No frame data available via API"

            return f"API extraction for '{data_type}' not implemented yet.\nAvailable: 'joints', 'frames'"

        except Exception as exc:
            return f"API extraction error: {exc}"

    def format_table_output(
        self,
        df: pd.DataFrame,
        format_type: str,
        output_file: str | None,
        description: str,
    ) -> str:
        if df is None or df.empty:
            return "No data to format"

        try:
            normalized = format_type.lower()
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")

            if normalized == "excel":
                file_path = output_file or f"etabs_table_{timestamp}.xlsx"
                df.to_excel(file_path, index=False)
                return f"Table exported to Excel: {file_path}\nRows: {len(df)} | Columns: {len(df.columns)}"

            if normalized == "csv":
                file_path = output_file or f"etabs_table_{timestamp}.csv"
                df.to_csv(file_path, index=False)
                return f"Table exported to CSV: {file_path}\nRows: {len(df)} | Columns: {len(df.columns)}"

            if normalized == "json":
                file_path = output_file or f"etabs_table_{timestamp}.json"
                df.to_json(file_path, orient="records", indent=2)
                return f"Table exported to JSON: {file_path}\nRows: {len(df)} | Columns: {len(df.columns)}"

            if normalized == "summary":
                return self.create_engineering_summary(df, description)

            return self.create_dataframe_preview(df)

        except Exception as exc:
            return f"Error formatting output: {exc}"

    def create_engineering_summary(self, df: pd.DataFrame, description: str) -> str:
        try:
            summary = "ENGINEERING SUMMARY\n"
            summary += f"Data: {len(df)} rows x {len(df.columns)} columns\n\n"

            numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns
            if len(numeric_cols) > 0:
                summary += "NUMERIC ANALYSIS:\n"
                for col in numeric_cols[:5]:
                    col_data = df[col].dropna()
                    if len(col_data) > 0:
                        summary += f"  {col}:\n"
                        summary += f"    Min: {col_data.min():.3f}\n"
                        summary += f"    Max: {col_data.max():.3f}\n"
                        summary += f"    Mean: {col_data.mean():.3f}\n"
                        if "ratio" in col.lower():
                            over_limit = (col_data > 1.0).sum()
                            if over_limit > 0:
                                summary += f"    {over_limit} elements over 1.0 limit\n"
                        summary += "\n"

            string_cols = df.select_dtypes(include=["object"]).columns
            if len(string_cols) > 0:
                summary += "CATEGORICAL DATA:\n"
                for col in string_cols[:3]:
                    unique_vals = df[col].value_counts().head(5)
                    summary += f"  {col}: {len(unique_vals)} unique values\n"
                    for val, count in unique_vals.items():
                        summary += f"    {val}: {count}\n"
                    summary += "\n"

            return summary

        except Exception as exc:
            return f"Error creating summary: {exc}"

    def create_dataframe_preview(self, df: pd.DataFrame) -> str:
        try:
            preview = "TABLE PREVIEW\n"
            preview += f"Shape: {len(df)} rows x {len(df.columns)} columns\n\n"
            preview += "COLUMNS:\n"
            for index, col in enumerate(df.columns):
                preview += f"  {index + 1}. {col} ({df[col].dtype})\n"

            preview += "\nFIRST 10 ROWS:\n"
            preview += df.head(10).to_string(max_cols=8, max_colwidth=12)
            if len(df) > 10:
                preview += f"\n\n... and {len(df) - 10} more rows"
            return preview

        except Exception as exc:
            return f"Error creating preview: {exc}"


_DEFAULT_EXTRACTOR = EtabsTableExtractor()


def etabs_extract(table_name: str) -> Optional[pd.DataFrame]:
    return _DEFAULT_EXTRACTOR.etabs_extract(table_name)


def get_etabs_table(table_description: str, export_format: str = "dataframe", output_file: str | None = None) -> str:
    return _DEFAULT_EXTRACTOR.get_etabs_table(table_description, export_format, output_file)


def find_table_name(description: str) -> Optional[str]:
    return _DEFAULT_EXTRACTOR.find_table_name(description)


def extract_joints_api() -> Optional[pd.DataFrame]:
    return _DEFAULT_EXTRACTOR.extract_joints_api()


def extract_frames_api() -> Optional[pd.DataFrame]:
    return _DEFAULT_EXTRACTOR.extract_frames_api()


def get_etabs_data_api(data_type: str) -> str:
    return _DEFAULT_EXTRACTOR.get_etabs_data_api(data_type)


def format_table_output(df: pd.DataFrame, format_type: str, output_file: str, description: str) -> str:
    return _DEFAULT_EXTRACTOR.format_table_output(df, format_type, output_file, description)


def create_engineering_summary(df: pd.DataFrame, description: str) -> str:
    return _DEFAULT_EXTRACTOR.create_engineering_summary(df, description)


def create_dataframe_preview(df: pd.DataFrame) -> str:
    return _DEFAULT_EXTRACTOR.create_dataframe_preview(df)


if __name__ == "__main__":
    print("ETABS TABLE EXTRACTION v1.0 - CLASS-BASED API")
    print("=" * 60)
    for example in ["frame forces", "joint displacements", "steel design", "material properties", "joint coordinates"]:
        print(f"  get_etabs_table('{example}')")
    print("\nAPI EXTRACTIONS:")
    print("  get_etabs_data_api('joints')")
    print("  get_etabs_data_api('frames')")

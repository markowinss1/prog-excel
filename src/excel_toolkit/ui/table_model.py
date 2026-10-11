"""Read-only rendering; values in the preview are never modified."""

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QPersistentModelIndex, Qt

from excel_toolkit.application.models import Preview

Index = QModelIndex | QPersistentModelIndex


class PreviewTableModel(QAbstractTableModel):
    def __init__(self) -> None:
        super().__init__()
        self.preview: Preview | None = None

    def set_preview(self, preview: Preview | None) -> None:
        self.beginResetModel()
        self.preview = preview
        self.endResetModel()

    def rowCount(self, parent: Index = QModelIndex()) -> int:  # noqa: B008
        return len(self.preview.rows) if self.preview and not parent.isValid() else 0

    def columnCount(self, parent: Index = QModelIndex()) -> int:  # noqa: B008
        return self.preview.column_count if self.preview and not parent.isValid() else 0

    def data(self, index: Index, role: int = Qt.ItemDataRole.DisplayRole) -> str | None:
        if not self.preview or not index.isValid():
            return None
        if index.row() >= len(self.preview.rows):
            return None
        row = self.preview.rows[index.row()]
        if index.column() >= len(row):
            return None
        cell = row[index.column()]
        value = "" if cell.value is None else str(cell.value)
        text = value[:2000] + ("…" if len(value) > 2000 else "")
        if role == Qt.ItemDataRole.DisplayRole:
            return text
        if role == Qt.ItemDataRole.ToolTipRole:
            return f"{text}\nТип: {cell.data_type}; формат: {cell.number_format}"
        return None

    def headerData(
        self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole
    ) -> str | None:
        if role == Qt.ItemDataRole.DisplayRole:
            return str(section + 1)
        return None

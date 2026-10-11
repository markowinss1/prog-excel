"""Read-only rendering; values in the preview are never modified."""

from PySide6.QtCore import (
    QAbstractListModel,
    QAbstractTableModel,
    QModelIndex,
    QPersistentModelIndex,
    Qt,
    Signal,
)

from excel_toolkit.application.models import Cell, Preview

Index = QModelIndex | QPersistentModelIndex


def column_letter(index: int) -> str:
    """Excel's stable column coordinate, independent of header text."""
    result = ""
    number = index + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


class ColumnSelectionModel(QAbstractListModel):
    """One checkable model entry per source column; no per-column widgets."""

    selection_changed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._headers: tuple[Cell, ...] = ()
        self._count = 0
        self._selected: set[int] = set()

    @property
    def selected_columns(self) -> tuple[int, ...]:
        return tuple(sorted(self._selected))

    def set_columns(
        self, headers: tuple[Cell, ...], count: int, *, preserve_selection: bool = False
    ) -> None:
        self.beginResetModel()
        self._headers = headers
        self._count = count
        self._selected = (
            {index for index in self._selected if index < count} if preserve_selection else set()
        )
        self.endResetModel()

    def column_label(self, index: int) -> str:
        letter = column_letter(index)
        value = self._headers[index].value if index < len(self._headers) else None
        header = str(value).strip() if value is not None else ""
        if not header:
            return f"Column {letter}"
        header = header[:300] + ("…" if len(header) > 300 else "")
        return f"{letter} — {header}"

    def rowCount(self, parent: Index = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else self._count

    def data(self, index: Index, role: int = Qt.ItemDataRole.DisplayRole) -> object:
        if not index.isValid() or not 0 <= index.row() < self._count:
            return None
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return self.column_label(index.row())
        if role == Qt.ItemDataRole.CheckStateRole:
            return (
                Qt.CheckState.Checked if index.row() in self._selected else Qt.CheckState.Unchecked
            )
        if role == Qt.ItemDataRole.UserRole:
            return index.row()
        return None

    def flags(self, index: Index) -> Qt.ItemFlag:
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable

    def setData(self, index: Index, value: object, role: int = Qt.ItemDataRole.EditRole) -> bool:
        if (
            role != Qt.ItemDataRole.CheckStateRole
            or not index.isValid()
            or not 0 <= index.row() < self._count
        ):
            return False
        checked = value in (Qt.CheckState.Checked, Qt.CheckState.Checked.value)
        if checked == (index.row() in self._selected):
            return False
        if checked:
            self._selected.add(index.row())
        else:
            self._selected.discard(index.row())
        self.dataChanged.emit(index, index, [Qt.ItemDataRole.CheckStateRole])
        self.selection_changed.emit()
        return True


class IssuesTableModel(QAbstractTableModel):
    """Render already-computed, bounded issue details without rerunning the engine."""

    def __init__(self, columns: ColumnSelectionModel) -> None:
        super().__init__()
        self._columns = columns
        self._issues: tuple = ()

    def set_issues(self, issues: tuple) -> None:
        self.beginResetModel()
        self._issues = issues
        self.endResetModel()

    def rowCount(self, parent: Index = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else len(self._issues)

    def columnCount(self, parent: Index = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else 3

    def data(self, index: Index, role: int = Qt.ItemDataRole.DisplayRole) -> str | None:
        if (
            role != Qt.ItemDataRole.DisplayRole
            or not index.isValid()
            or not 0 <= index.row() < len(self._issues)
            or not 0 <= index.column() < 3
        ):
            return None
        issue = self._issues[index.row()]
        return (
            str(issue.source_row),
            self._columns.column_label(issue.column - 1),
            issue.kind,
        )[index.column()]

    def headerData(
        self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole
    ) -> str | None:
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return ("Source Row", "Column", "Problem")[section]
            return str(section + 1)
        return None


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

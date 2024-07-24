import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFormLayout, QMessageBox, QTreeView
from PyQt5.QtGui import QStandardItemModel, QStandardItem

class PenduleApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Pendule Calculator')
        layout = QVBoxLayout()

        self.formLayout = QFormLayout()
        self.nInput = QLineEdit(self)
        self.eInput = QLineEdit(self)
        self.sheetInput = QLineEdit(self)

        self.formLayout.addRow(QLabel('N:'), self.nInput)
        self.formLayout.addRow(QLabel('e:'), self.eInput)
        self.formLayout.addRow(QLabel('Sheet Name:'), self.sheetInput)

        self.calculateButton = QPushButton('Calculate', self)
        self.calculateButton.clicked.connect(self.calculatePositions)
        self.nextButton = QPushButton('Suivant', self)
        self.nextButton.clicked.connect(self.checkSum)
        self.nextButton.setEnabled(False)

        self.treeView = QTreeView(self)
        self.model = QStandardItemModel()
        self.treeView.setModel(self.model)

        layout.addLayout(self.formLayout)
        layout.addWidget(self.calculateButton)
        layout.addWidget(self.nextButton)
        layout.addWidget(self.treeView)

        self.setLayout(layout)

    def calculatePositions(self):
        try:
            N = float(self.nInput.text())
            e = float(self.eInput.text())
            sheet_name = self.sheetInput.text().strip()
        except ValueError:
            QMessageBox.warning(self, 'Input Error', 'Please enter valid numbers for N and e.')
            return

        try:
            df = pd.read_excel('pendules.xlsx', sheet_name=sheet_name)
            df = df.dropna(subset=['N', 'e'])
        except ValueError as ve:
            QMessageBox.critical(self, 'Sheet Error', f'Error with sheet name: {ve}')
            return
        except Exception as e:
            QMessageBox.critical(self, 'File Error', f'Error loading Excel file: {e}')
            return

        filtered_df = df[(df['N'] == N) & (df['e'] == e)]
        if not filtered_df.empty:
            distance_cols = [col for col in filtered_df.columns if col.startswith('D')]
            if not distance_cols:
                QMessageBox.warning(self, 'Data Error', 'No distance columns found in the data.')
                return

            distances = filtered_df[distance_cols].values.flatten()
            distances = [d for d in distances if pd.notna(d)]
        else:
            closest_below = df[df['N'] < N].sort_values(by='N', ascending=False).head(1)
            closest_above = df[df['N'] > N].sort_values(by='N').head(1)
            if closest_below.empty or closest_above.empty:
                QMessageBox.warning(self, 'Data Error', 'No data found for the provided N and e, and cannot interpolate.')
                return

            closest_below_N = closest_below['N'].values[0]
            closest_above_N = closest_above['N'].values[0]
            closest_below_distances = closest_below.filter(like='D').values.flatten()
            closest_above_distances = closest_above.filter(like='D').values.flatten()
            interpolated_distances = closest_below_distances + (closest_above_distances - closest_below_distances) * (N - closest_below_N) / (closest_above_N - closest_below_N)
            distances = interpolated_distances

            total_distance = sum(distances)
            if total_distance != N:
                D1 = distances[0]
                remaining_distance = N - D1
                remaining_distances = distances[1:]
                remaining_total = sum(remaining_distances)
                if remaining_total != 0:
                    scale_factor = remaining_distance / remaining_total
                    adjusted_distances = [D1] + [round(d * scale_factor, 2) for d in remaining_distances]
                else:
                    QMessageBox.warning(self, 'Data Error', 'Remaining distances total is zero, cannot scale distances.')
                    return
            else:
                adjusted_distances = distances

            def redistribute(distances, max_value):
                excess_total = sum(d - max_value for d in distances if d > max_value)
                distances = [min(d, max_value) for d in distances]
                redistribute_amount = excess_total / len(distances)
                distances = [round(d + redistribute_amount, 2) for d in distances]
                return distances

            adjusted_distances = redistribute(adjusted_distances, 8)

            total_adjusted = sum(adjusted_distances)
            if total_adjusted != N:
                adjusted_distances[-1] = round(adjusted_distances[-1] + (N - total_adjusted), 2)
            distances = adjusted_distances

        # Combiner les petites distances avec les précédentes
        combined_distances = []
        for distance in distances:
            if combined_distances and distance < 2:
                combined_distances[-1] += distance
            else:
                combined_distances.append(distance)
        
        # Limiter le nombre de distances à moins de 9
        while len(combined_distances) > 9:
            min_distance = min(combined_distances)
            min_index = combined_distances.index(min_distance)
            if min_index > 0:
                combined_distances[min_index - 1] += combined_distances.pop(min_index)
            else:
                combined_distances[1] += combined_distances.pop(0)

        # Arrondir les distances à 0.25 près
        combined_distances = [round(d * 4) / 4 for d in combined_distances]

        total_combined = sum(combined_distances)
        if total_combined != N:
            combined_distances[-1] = round(combined_distances[-1] + (N - total_combined), 2)
        
        distances = combined_distances

        if len(distances) < 2:
            QMessageBox.warning(self, 'Data Error', 'Not enough data to calculate positions.')
            return

        positions = [0]
        for i in range(1, len(distances) + 1):
            positions.append(round(positions[-1] + distances[i - 1], 2))

        self.model.clear()
        self.model.setHorizontalHeaderLabels(['Type', 'Value'])

        distances_item = QStandardItem('Distances')
        self.model.appendRow(distances_item)
        for i, d in enumerate(distances, start=1):
            distances_item.appendRow([QStandardItem(f'Distance {i}'), QStandardItem(f'{d:.2f}')])

        positions_item = QStandardItem('Positions')
        self.model.appendRow(positions_item)
        for i, p in enumerate(positions, start=1):
            positions_item.appendRow([QStandardItem(f'Position {i}'), QStandardItem(f'{p:.2f}')])

        self.treeView.expandAll()
        self.nextButton.setEnabled(True)

    def checkSum(self):
        distances = []
        for i in range(self.model.item(0).rowCount()):
            distance_item = self.model.item(0).child(i, 1)
            distances.append(float(distance_item.text()))
        total_distance = sum(distances)
        N = float(self.nInput.text())
        if abs(total_distance - N) < 0.05:
            QMessageBox.information(self, 'Verification', f'The sum of distances is {total_distance}, which is approximately equal to N ({N}).')
        else:
            QMessageBox.warning(self, 'Verification', f'The sum of distances is {total_distance}, which is not equal to N ({N}).')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PenduleApp()
    ex.resize(400, 300)
    ex.show()
    sys.exit(app.exec_())
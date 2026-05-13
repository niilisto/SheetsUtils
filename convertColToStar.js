function convertColumnsToStars() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const targetColumns = ["B", "D", "F"];
  const lastRow = sheet.getLastRow();

  targetColumns.forEach(letter => {
    const range = sheet.getRange(`${letter}1:${letter}${lastRow}`);
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      const rating = values[i][0];
      if (typeof rating === 'number' && rating >= 0 && rating <= 5) {
        const stars = "★".repeat(rating) + "☆".repeat(5 - rating);
        values[i][0] = stars;
      }
    }

    range.setValues(values);
  });
}
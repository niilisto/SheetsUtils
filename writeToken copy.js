function onlyRunOnColumns(e, targetColumns, callback) {
  if (!e || !e.range) return;

  const editedCol = e.range.getColumn();
  const targetColsNum = targetColumns.map(letter => letterCol(letter));

  if (targetColsNum.includes(editedCol)) {
    try {
      callback();
    } catch (err) {
      Logger.log(`Erro em onlyRunOnColumns para colunas [${targetColumns.join(',')}]: ${err}`);
    }
  }
}

function letterCol(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + letter.charCodeAt(i) - 64;
  }
  return column;
}

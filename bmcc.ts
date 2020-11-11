import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

enum TAB_NAMES {
  BOARD = "Board",
  PLAYER = "Player",
  QUESTIONS = "Questions",
  ANSWERS = "Answers",
}

// TODO allow to "continue" a game (just reset scores and keep used answers)
// FIXME update layout
// FIXME Allow to refuse a question
// FIXME choose answer by selection only
// FIXME show only "new question" for questioner
// FIXME question not updated for one of us
// FIXME show only "validate vote" when "time to chose answer"
// FIXME time to update question (and to get new cards) but not to get it (the popup comes fast)
// disable edition of question / answers

const translations = {
  "fr": {}
};

class MathUtils {
  static getRandomInt(max: number): number {
    return Math.floor(Math.random() * Math.floor(max));
  }

  static getRandomIntButNot(max: number, except: number[]) {
    let number;
    do {
      number = MathUtils.getRandomInt(max);
    } while (except.includes(number));

    return number;
  }
}

class Ui {
  static alert(message: string) {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(message, ui.ButtonSet.OK);
  }

  static confirm(message: string): boolean {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(message, ui.ButtonSet.OK_CANCEL) === ui.Button.OK;
  }
}

abstract class Tab {
  constructor(public sheet: Sheet) {
  }
}

abstract class PlayerTab extends Tab {
  answerNameCell: GoogleAppsScript.Spreadsheet.Range;
  answerCell: GoogleAppsScript.Spreadsheet.Range;
  cardCell: GoogleAppsScript.Spreadsheet.Range;
  questionCell: GoogleAppsScript.Spreadsheet.Range;
  messageCell: GoogleAppsScript.Spreadsheet.Range;
  hallOfFameCell: GoogleAppsScript.Spreadsheet.Range;
}

class TemplatePlayerTab extends PlayerTab {
  private findCell(name: string): GoogleAppsScript.Spreadsheet.Range {
    const cells = this.tab.getRange("A1:G50").getValues();
    for (let rowNum = 0; rowNum < cells.length; rowNum++) {
      const row = cells[rowNum];
      for (let colNum = 0; colNum < row.length; colNum++) {
        if (row[colNum] == name) {
          return this.tab.getRange(rowNum + 1, colNum + 1);
        }
      }
    }
  }

  constructor(public tab: Sheet) {
    super(tab);
    this.answerNameCell = this.findCell('<name>');
    this.answerCell = this.findCell('<answer>');
    this.questionCell = this.findCell('<question>');
    this.cardCell = this.findCell('<my_card>');
    this.messageCell = this.findCell('<message>');
    this.hallOfFameCell = this.findCell('Hall of fame').offset(1, 0);
  }
}

class RealPlayerTab extends PlayerTab {
  constructor(public tab: GoogleAppsScript.Spreadsheet.Sheet, public initialCardCount: number, template: TemplatePlayerTab, private cardsCount: number) {
    super(tab);
    this.answerNameCell = this.tab.getRange(template.answerNameCell.getA1Notation());
    this.answerCell = this.tab.getRange(template.answerCell.getA1Notation());
    this.questionCell = this.tab.getRange(template.questionCell.getA1Notation());
    this.cardCell = this.tab.getRange(template.cardCell.getA1Notation());
    this.messageCell = this.tab.getRange(template.messageCell.getA1Notation());
    this.hallOfFameCell = this.tab.getRange(template.hallOfFameCell.getA1Notation());
  }

  setMessage(message: string, color?: string) {
    this.messageCell.setValue(message);
    this.messageCell.setBackground(color);
  }

  copyRows(fromRow: number, pasteCount: number) {
    const questionRange = this.tab.getRange(`${fromRow}:${fromRow}`);
    this.tab.insertRowsAfter(fromRow, pasteCount);
    const nextRange = this.tab.getRange(`${fromRow + 1}:${fromRow + pasteCount}`);
    questionRange.copyTo(nextRange);
  };

  init(answerersCount: number) {
    // Important! First card because we know it comes _after_ question
    // Copy card row for as many as initial cards
    this.copyRows(this.cardCell.getRow(), this.initialCardCount - 1);

    // Copy answers row for as many as "answerers" there are
    const answerRow = this.answerNameCell.getRow();
    this.copyRows(answerRow, answerersCount - 1);

    this.updateCells(this.initialCardCount, answerersCount);

    // Add fore scoreboard
    this.copyRows(this.hallOfFameCell.getRow(), answerersCount);
  }

  private addRows(range: GoogleAppsScript.Spreadsheet.Range, rowsToAdd: number) {
    return range.offset(rowsToAdd, 0);
  }

  updateCells(initCardCount: number, answerersCount: number) {
    // TODO Because we know everything comes after hall of fame and cards are _after_ question
    this.messageCell = this.addRows(this.messageCell, answerersCount);
    this.answerNameCell = this.addRows(this.answerNameCell, answerersCount);
    this.answerCell = this.addRows(this.answerCell, answerersCount);
    this.questionCell = this.addRows(this.questionCell, answerersCount);
    this.cardCell = this.addRows(this.cardCell, answerersCount + answerersCount - 1);
    return this;
  }

  getCardsRange() {
    return this.sheet.getRange(
      this.cardCell.getRow(),
      this.cardCell.getColumn(),
      this.initialCardCount,
      2 // TODO hardcoded
    )
  }

  get question(): string {
    return this.questionCell.getValue() as string;
  }

  getSelectedCardsRanges(): GoogleAppsScript.Spreadsheet.Range[] {
    const cards = this.getCardsRange();
    const selectedCards: GoogleAppsScript.Spreadsheet.Range[] = [];
    cards.getValues().forEach((row, i) => {
      const selection = row[1];
      if (selection === 1) {
        selectedCards[0] = this.sheet.getRange(cards.getRow() + i, cards.getColumn(), 1, 2);
      }
      if (selection === 2) {
        selectedCards[1] = this.sheet.getRange(cards.getRow() + i, cards.getColumn(), 1, 2);
      }
      if (selection === 3) {
        selectedCards[2] = this.sheet.getRange(cards.getRow() + i, cards.getColumn(), 1, 2);
      }
    });

    return selectedCards;
  }

  getAnswers(): GoogleAppsScript.Spreadsheet.Range {
    // TODO hardcoded numColumns
    return this.tab.getRange(this.answerNameCell.getRow(), this.answerNameCell.getColumn(), this.cardsCount, 3);
  }
}

class BoardPlayer {
  nameCell: GoogleAppsScript.Spreadsheet.Range;
  scoreCell: GoogleAppsScript.Spreadsheet.Range;
  answerCell: GoogleAppsScript.Spreadsheet.Range;
  randomCell: GoogleAppsScript.Spreadsheet.Range;
}

class BoardTab extends Tab {
  questionerCell: GoogleAppsScript.Spreadsheet.Range;
  answerCountCell: GoogleAppsScript.Spreadsheet.Range;
  nameStartCell: GoogleAppsScript.Spreadsheet.Range;
  players: { [name: string]: BoardPlayer } = {};
  playerNames: string[];

  // TODO make it guessed from tab
  private nameColumn = 1;
  private scoreColumn = 2;
  private answerColumn = 3;
  private randomColumn = 4;

  constructor(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    super(sheet);
    // TODO make it guessed from tab
    this.questionerCell = this.sheet.getRange("B1");
    this.answerCountCell = this.sheet.getRange("C1");
    this.nameStartCell = this.sheet.getRange("A2");

    // TODO iterate through cells instead
    this.range.getValues().forEach((row: string[], i: number) => {
      const rowNum = this.nameStartCell.getRow() + 1 + i;
      const name = row[0];
      this.players[name] = {
        nameCell: this.sheet.getRange(rowNum, this.nameColumn),
        scoreCell: this.sheet.getRange(rowNum, this.scoreColumn),
        answerCell: this.sheet.getRange(rowNum, this.answerColumn),
        randomCell: this.sheet.getRange(rowNum, this.randomColumn),
      }
    });
    this.playerNames = Object.keys(this.players);
  }

  private get range(): GoogleAppsScript.Spreadsheet.Range {
    const fullRange = this.nameStartCell.getFilter().getRange();
    return fullRange.offset(1, 0, fullRange.getNumRows() - 1);
  }

  resetAnswers() {
    const range = this.range;
    const values = range.getValues();
    values.map(row => {
      // TODO works because table is in first row
      row[this.answerColumn - 1] = '';
      row[this.randomColumn - 1] = '';
    });

    range.setValues(values);
  }

  get usedRandoms(): number[] {
    const allRange = this.nameStartCell.getFilter().getRange();
    const randomRange = this.sheet.getRange(
      allRange.getRow() + 1,
      this.randomColumn,
      this.playerNames.length
    );
    return randomRange.getValues().map(row => row[0]).filter(v => v !== "") as number[];
  }

  get answersCount(): number {
    return this.answerCountCell.getValue() as number;
  }

  get currentQuestioner(): string {
    return this.questionerCell.getValue() as string;
  }
}

class CardsTab extends Tab {
  availableColumn = 2;
  textColumn = 1;
  columns = [this.textColumn, this.availableColumn];
  firstRow = 2;
  cardsCountCell = this.sheet.getRange(1, 2);
  availableCardsCountCell = this.sheet.getRange(1, 3);


  private get range(): GoogleAppsScript.Spreadsheet.Range {
    const fullRange = this.sheet.getRange(this.firstRow, Math.min(...this.columns), this.totalCardsCount, this.columns.length).getFilter().getRange();
    return fullRange.offset(1, 0, fullRange.getNumRows() - 1);
  }

  getNewCards(cardsCount: number): string[] {
    const availableCardsCount = this.availableCardsCountCell.getValue() as number;

    if (cardsCount > availableCardsCount) {
      Ui.alert("Not enough questions left!");
      throw "error";
    }

    const allCards = this.range;

    const allCardsValues = allCards.getValues();
    const selectedCards = [];

    while (selectedCards.length < cardsCount) {
      const randomRow = this.firstRow + MathUtils.getRandomInt(availableCardsCount);
      // TODO only works because table starts at first column
      if (!allCardsValues[randomRow][this.availableColumn - 1] as boolean) {
        continue;
      }
      allCardsValues[randomRow][this.availableColumn - 1] = false;
      selectedCards.push(allCardsValues[randomRow][this.textColumn - 1]);
    }

    allCards.setValues(allCardsValues);

    return selectedCards;
  }

  get totalCardsCount() {
    return this.cardsCountCell.getValue() as number;
  }

  resetUsedCards() {
    const range = this.range;
    const numRows = range.getNumRows();
    for (let row = 1; row <= numRows; row++) {
      range.getCell(row, this.availableColumn).setValue(true);
    }
  }
}

class Bmcc {
  doc: Spreadsheet;

  initCardCount = 10;
  minimumCardCount = 5;

  private readonly questionTab: CardsTab;
  private readonly answersTab: CardsTab;
  private readonly playerTemplateTab: TemplatePlayerTab;
  private readonly boardTab: BoardTab;

  constructor() {
    this.doc = SpreadsheetApp.getActiveSpreadsheet();
    this.boardTab = new BoardTab(this.getTab(TAB_NAMES.BOARD));
    this.questionTab = new CardsTab(this.getTab(TAB_NAMES.QUESTIONS));
    this.answersTab = new CardsTab(this.getTab(TAB_NAMES.ANSWERS));
    this.playerTemplateTab = new TemplatePlayerTab(this.getTab(TAB_NAMES.PLAYER));
  }

  private static playerTabs: {[name: string]: RealPlayerTab} = {};

  getPlayerTab(playerName: string): RealPlayerTab {
    if (Bmcc.playerTabs[playerName] === undefined) {
      // Keep in cache
      Bmcc.playerTabs[playerName] = new RealPlayerTab(this.getTab(playerName), this.initCardCount, this.playerTemplateTab, this.answerersCount)
        .updateCells(this.initCardCount, this.answerersCount);
    }

    return Bmcc.playerTabs[playerName];
  }

  get answerersCount() {
    return this.boardTab.playerNames.length - 1;
  }

  createBoard() {
    // remove existing players
    this.removePlayerTabs();

    this.resetScores();

    // reset used questions and answers
    this.questionTab.resetUsedCards();
    this.answersTab.resetUsedCards();

    const playerNames = this.boardTab.playerNames;

    const randomQuestioner = MathUtils.getRandomInt(playerNames.length - 1);

    playerNames.forEach(player => {
      const tab = new RealPlayerTab(this.copyTab(this.playerTemplateTab.tab.getName(), player), this.initCardCount, this.playerTemplateTab, this.answerersCount);

      tab.init(this.answerersCount);

      // TODO add names for board?
      this.completeCards(tab, this.initCardCount);

      tab.setMessage('');
      tab.questionCell.setValue('');
    });

    this.cleanRound();

    Logger.log(`random player: ${randomQuestioner}`);

    const questionerName = playerNames[randomQuestioner];

    Ui.alert(`${questionerName} a été tiré au sort pour commencer !`)

    this.setQuestioner(questionerName);
  }

  resetScores() {
    this.boardTab.playerNames.forEach(name => {
      this.boardTab.players[name].scoreCell.setValue(0);
    });
  }

  takeNewQuestion() {
    const currentUser = this.doc.getActiveSheet().getName();

    // check is questioner
    if (this.boardTab.questionerCell.getValue() !== currentUser) {
      return;
    }

    const question = this.questionTab.getNewCards(1)[0];

    this.cleanRound();

    // Update question everywhere
    this.boardTab.playerNames.forEach(p => {
      const playerTab = this.getPlayerTab(p);
      playerTab.questionCell.setValue(question);
      if (p === currentUser) {
        playerTab.setMessage('En attente des propositions...');
      } else {
        playerTab.setMessage("C'est le moment de choisir ta meilleure proposition !", '#daedf5');
      }
    });

    Ui.alert(question);
  }

  cleanRound() {
    // Clean board
    this.boardTab.resetAnswers();

    // Clean players tab
    this.boardTab.playerNames.forEach(p => {
      // Reset question
      const playerTab = this.getPlayerTab(p);

      // Reset answers
      const answers = playerTab.getAnswers();
      let emptyRow = new Array(answers.getNumColumns());
      emptyRow = emptyRow.fill('');
      let newValues = new Array(answers.getNumRows());
      newValues = newValues.fill(emptyRow);

      answers.setValues(newValues);

      this.completeCards(playerTab, this.minimumCardCount);
    })
  }

  forEachCell(range: GoogleAppsScript.Spreadsheet.Range,
              callback: (cell: GoogleAppsScript.Spreadsheet.Range, row: number, column: number) => any) {
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    for (let row = 1; row <= numRows; row++) {
      for (let col = 1; col <= numCols; col++) {
        callback(range.getCell(row, col), row, col);
      }
    }
  }

  private removePlayerTabs() {
    const sheets = this.doc.getSheets();

    const names = Object.keys(TAB_NAMES).map(k => TAB_NAMES[k as any]);

    for (const iSheet in sheets) {
      const name = sheets[iSheet].getName();
      if (names.includes(name))
        continue;
      this.doc.deleteSheet(sheets[iSheet]);
    }
  }

  private getTab(tabName: string): Sheet | undefined {
    const sheets = this.doc.getSheets();

    for (const iSheet in sheets) {
      if (tabName === sheets[iSheet].getName()) {
        return sheets[iSheet];
      }
    }

    return undefined;
  }

  private copyTab(fromName: string, toName: string, position?: number): Sheet {
    const from = this.getTab(fromName);
    const to = from.copyTo(this.doc);
    to.setName(toName);
    to.activate();

    if (position) {
      this.doc.moveActiveSheet(position);
    }

    return to;
  }

  private completeCards(tab: RealPlayerTab, upToCount: number) {
    let emptyCells: GoogleAppsScript.Spreadsheet.Range[] = [];
    this.forEachCell(tab.getCardsRange(),
      (cell, row, column) => {
        switch (column) {
          // Reset choice
          case 2: // TODO hardcoded column number
            cell.setValue('');
            break;
          // Register card value
          case 1:
            // TODO <my_card>
            if (cell.getValue() === '' || cell.getValue() === '<my_card>') {
              emptyCells.push(cell);
            }
            break;
        }
      });

    const cardsToAdd = upToCount - (this.initCardCount - emptyCells.length);

    if (cardsToAdd > 0) {
      // To take first the first rows
      emptyCells = emptyCells.reverse();

      const cards = this.answersTab.getNewCards(cardsToAdd);

      for (let i = 0; i < cardsToAdd; i++) {
        const cardText = cards.pop();
        const cell = emptyCells.pop();
        cell.setValue(cardText);
      }
    }
  }

  registerCard() {
    const currentUser = this.doc.getActiveSheet().getName();

    // check is not questioner
    const questioner = this.boardTab.questionerCell.getValue();
    if (questioner === currentUser) {
      return;
    }

    const playerTab = this.getPlayerTab(currentUser);

    const selectedCards = playerTab.getSelectedCardsRanges();
    const question = playerTab.question;
    let fullAnswer = question;

    const replacements = ['XXX', 'YYY', 'ZZZ'];
    selectedCards.forEach((range, i) => {
      let text = selectedCards[i].getCell(1, 1).getValue() as string;
      text = text.charAt(0).toLowerCase() + text.slice(1);
      fullAnswer = fullAnswer.replace(replacements[i], text);
    });

    if (fullAnswer.includes('XXX')) {
      Ui.alert('Choisis au moins une réponse !');
      return;
    }
    if (fullAnswer.includes('YYY')) {
      Ui.alert('Choisis deux réponses !');
      return;
    }
    if (fullAnswer.includes('ZZZ')) {
      Ui.alert('Choisis trois réponses !');
      return;
    }

    if (!Ui.confirm(`"${fullAnswer}"\nTu confirmes ?`)) {
      return;
    }

    playerTab.sheet.setTabColor('#210cff');

    this.boardTab.players[currentUser].answerCell.setValue(fullAnswer);

    // for case where user has already replied
    this.boardTab.players[currentUser].randomCell.setValue('');
    this.boardTab.players[currentUser].randomCell.setValue(MathUtils.getRandomIntButNot(this.answerersCount, this.boardTab.usedRandoms));

    if (this.boardTab.answersCount == this.answerersCount) {
      this.endSelectionPhase();
    }
  }

  endSelectionPhase() {
    const questionerTab = this.getPlayerTab(this.boardTab.currentQuestioner);
    questionerTab.setMessage("C'est le moment de choisir la meilleure combinaison !", '#c8f7ca');

    this.boardTab.playerNames
      .filter(name => name !== this.boardTab.currentQuestioner)
      .forEach(name => {
        const player = this.boardTab.players[name];
        const answer = player.answerCell.getValue() as string;
        const random = player.randomCell.getValue() as number;
        const answerInQuestionerTab = questionerTab.sheet.getRange(
          questionerTab.answerCell.getRow() + random,
          questionerTab.answerCell.getColumn()
        );
        answerInQuestionerTab.setValue(answer);

        // remove card from player's tab
        const playerTab = this.getPlayerTab(name);
        const selectedCards = playerTab.getSelectedCardsRanges();
        selectedCards.forEach(range => {
          range.setValues([['', '']]); // TODO because we know there are only 2 columns
        });
        playerTab.setMessage('');
      });
  }

  registerVote() {
    const currentUser = this.doc.getActiveSheet().getName();

    // check is questioner
    const currentQuestioner = this.boardTab.questionerCell.getValue() as string;
    if (currentQuestioner !== currentUser) {
      return;
    }

    const questionerTab = this.getPlayerTab(this.boardTab.currentQuestioner);
    const answers = questionerTab.getAnswers();
    const answerValues = answers.getValues();

    // TODO hardcoded column
    const chosen = answerValues.filter(row => row[2] !== undefined && row[2] !== "");

    // check one and only one vote
    if (chosen.length !== 1) {
      Ui.alert('Hey ! Il faut choisir UNE réponse');
      return;
    }

    const chosenRow = chosen[0];
    const answer = chosenRow[1] as string;

    this.boardTab.playerNames.forEach(name => {
      const player = this.boardTab.players[name];
      this.getPlayerTab(name).sheet.setTabColor(null);
      for (let i = 0; i < answerValues.length; i++) {
        const row = answerValues[i];
        // TODO hardcoded column
        if (row[1] === player.answerCell.getValue()) {
          questionerTab.sheet.getRange(
            answers.getRow() + i,
            // TODO hardcoded column
            answers.getColumn()
          ).setValue(name);

          if (row[2] === 1) {
            // Winner!
            player.scoreCell.setValue(player.scoreCell.getValue() as number + 1);
            Ui.alert(`And the winner is... ${name}!`);

            // FIXME handle end of game

            // Reset questioner tab
            const playerTab = this.getPlayerTab(currentQuestioner);
            playerTab.setMessage('');

            this.setQuestioner(name);
          }
        }
      }
    });

  }

  private setQuestioner(name: string) {
    // update new questioner in board
    this.boardTab.questionerCell.setValue(name);

    // change questioner's tab color
    const playerTab = this.getPlayerTab(name);

    const drawings = playerTab.sheet.getDrawings();
    const drawing = drawings.filter((e) => e.getOnAction() == 'takeNewQuestion')[0];
    drawing.setWidth(200);
    drawing.setHeight(10);

    playerTab.sheet.setTabColor('red');
    playerTab.setMessage('Choisis une nouvelle question !', '#f5d0d0');
  }
}

const bmcc = new Bmcc();

function start() {
  bmcc.createBoard();
}

function resetScores() {
  bmcc.resetScores();
}

function takeNewQuestion() {
  bmcc.takeNewQuestion();
}

function registerCard() {
  bmcc.registerCard();
}

function registerVote() {
  bmcc.registerVote();
}

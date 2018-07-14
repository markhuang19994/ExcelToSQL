const { Component } = React;
const myMurmur = `Please use chrome, I don't want to consider other browser`;
class App extends Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      template: {
        table: "",
        columns: "",
        defaultInsertTempLate: `INSERT INTO {table} (OID,{columns}) VALUES(REPLACE(NEWID(), '-', ''),{values});`,
        insertTemplate: this.defaultInsertTempLate
      },
      dragFile: null
    };
    this.setDragFile = this.setDragFile.bind(this);
    this.tableHandler = this.tableHandler.bind(this);
    this.columnHandler = this.columnHandler.bind(this);
  }

  setDragFile(dragFile) {
    this.state.dragFile = dragFile;
    this.setState(this.state);
    console.log("file name = " + this.state.dragFile.name);
  }

  tableHandler(ev) {
    this.state.template.insertTemplate = this.state.template.defaultInsertTempLate
      .replace("{table}", ev.target.value)
      .replace("{columns}", this.state.template.columns.toUpperCase());
    this.state.template.table = ev.target.value;
    this.setState(this.state);
  }

  columnHandler(ev) {
    let columns = ev.target.value.toUpperCase().replace(/ /g, "");
    this.state.template.insertTemplate = this.state.template.defaultInsertTempLate
      .replace("{table}", this.state.template.table.toUpperCase())
      .replace("{columns}", columns);
    this.state.template.columns = columns;
    this.setState(this.state);
  }

  render() {
    return (
      <div className="App">
        <section>
          <h1 className="main-title">Convert excel to SQL</h1>
          <DropFile setDragFile={dragFile => this.setDragFile(dragFile)}>
            Drag your excel file to zone
          </DropFile>
        </section>
        <section>
          <div className="main-div main_excel_parameter">
            <p>
              Your File : {this.state.dragFile ? this.state.dragFile.name : ""}
            </p>
            <label htmlFor="table">Enter table scema :&nbsp;</label>
            <input
              type="text"
              id="table"
              name="table"
              placeholder="[dbo].[PCL_MEDIA]"
              onChange={this.tableHandler}
            />
            <br />
            <label htmlFor="columns">Enter table columns :&nbsp;</label>
            <input
              type="text"
              id="columns"
              name="columns"
              placeholder="ECID_ICID, MEDIA, CHANNEL, SEGMENT ,SOURCECODE_PCL ,SOURCECODE_UPL"
              onChange={this.columnHandler}
            />
            <br />
            <button
              onClick={() => {
                generateSql(
                  this.state.dragFile,
                  this.state.template.insertTemplate,
                  this.state.template.columns.split(",").length + 1
                );
              }}
            >
              Convert
            </button>
          </div>
        </section>
      </div>
    );
  }
}

class DropFile extends Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      dragFile: null
    };
    this.dropHandler = this.dropHandler.bind(this);
  }

  dropHandler(ev) {
    ev.preventDefault();
    if (ev.dataTransfer.items) {
      if (
        ev.dataTransfer.items.length === 1 &&
        ev.dataTransfer.items[0].kind === "file"
      ) {
        this.state.dragFile = ev.dataTransfer.items[0].getAsFile();
        this.setState(this.state);
        this.props.setDragFile(this.state.dragFile);
      }
    } else {
      console.error(myMurmur);
    }
    this.removeDragData(ev);
  }

  dragOverHandler(ev) {
    ev.preventDefault();
  }

  removeDragData(ev) {
    ev.dataTransfer.items
      ? ev.dataTransfer.items.clear()
      : console.error(myMurmur);
  }

  render() {
    return (
      <div className="main-drop">
        <div
          id="drop_zone"
          className="main-drop_container div-text-vertical-center main-div"
          onDrop={this.dropHandler}
          onDragOver={this.dragOverHandler}
        >
          <p className="main-drop_text">{this.props.children}</p>
        </div>
      </div>
    );
  }
}

ReactDOM.render(<App />, document.getElementById("root"));

function generateSql(dragFile, insertTemplate, limit) {
  dragFile && new ExcelToJSON().parseExcel(dragFile, insertTemplate, limit);
}

class ExcelToJSON {
  constructor() {
    this.parseExcel = function(file, insertTemplate = "", limit = 0) {
      let reader = new FileReader();
      reader.onload = function(e) {
        let data = e.target.result;
        let workbook = XLSX.read(data, {
          type: "binary"
        });
        workbook.SheetNames.forEach(function(sheetName) {
          let XL_row_object = XLSX.utils.sheet_to_row_object_array(
            workbook.Sheets[sheetName]
          );
          let allInsertSql = "";
          XL_row_object.forEach(row => {
            let max = limit;
            let i = 0;
            let cell = [];
            let insertSql = "";
            Object.keys(row).forEach(column => {
              i < max && cell.push(row[column]);
              i++;
            });
            insertSql = insertTemplate.replace(
              "{values}",
              `'${cell.map(x => x.replace("'", "''")).join("','")}'`
            );
            allInsertSql += insertSql + "\r\n";
          });
          download("result.sql", allInsertSql);
        });
      };
      reader.onerror = function(ex) {
        console.log(ex);
      };
      reader.readAsBinaryString(file);
    };
  }
}

function download(filename, text) {
  var e = document.createElement("a");
  e.setAttribute(
    "href",
    "data:text/plain;charset=utf-8," + encodeURIComponent(text)
  );
  e.setAttribute("download", filename);
  e.style.display = "none";
  document.body.appendChild(e);
  e.click();
  document.body.removeChild(e);
}

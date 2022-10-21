import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global Word, require */

export interface AppProps {
  title?: string;
  isOfficeInitialized?: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  fontSize: number;
}

async function getCommentsData() {
  const resp = await fetch("https://jsonplaceholder.typicode.com/comments");
  const data = await resp.json();
  const commentData = data.map((coment) => {
    return coment.body;
  });
  return commentData;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      fontSize: 14,
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {
    return Word.run(async (context) => {
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
      paragraph.font.color = "blue";
      await context.sync();
    });
  };

  getTooXml = async () => {
    return Word.run(async (context) => {
      var body = context.document.body;
      var bodyOOXML = body.getOoxml();
      await context.sync();
      console.log("Body OOXML contents: ", bodyOOXML.value);
    });
  };

  getHtml = async () => {
    return Word.run(async (context) => {
      var body = context.document.body;

      var bodyHTML = body.getHtml();

      const range = context.document.getSelection();
      const html = range.getHtml();
      range.load("value");
      await context.sync();
      console.log(html, range);
    });
  };

  contentControl = async () => {
    return await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.clear();
      const wordContentControl = range.insertContentControl();
      wordContentControl.tag = "Name";
      wordContentControl.title = "Aashish Manani";
      wordContentControl.cannotEdit = false;
      wordContentControl.appearance = "BoundingBox";
      wordContentControl.font.color = "red";
      await context.sync();
    }).catch((error) => {
      console.log(error);
    });
  };

  Function = async () => {
    return await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.clear();
      const wordContentControl = range.insertContentControl();
      wordContentControl.tag = "Name";
      wordContentControl.title = "Aashish Manani";
      wordContentControl.cannotEdit = false;
      wordContentControl.appearance = "BoundingBox";
      wordContentControl.style = "Aashish";
      wordContentControl.font.color = "red";
      Word.ContentControlType.checkBox;
      await context.sync();
    });
  };

  getDataFromApi = async () => {
    const comments = await getCommentsData();
    return await Word.run(async (context) => {
      comments.forEach((element, index) => {
        const paragraph = context.document.body.insertParagraph(
          index + 1 + " ." + element.replace(/\n/g, " "),
          Word.InsertLocation.end
        );
        if (index % 2 === 0) {
          paragraph.font.color = "black";
        } else {
          paragraph.font.color = "blue";
        }
        paragraph.font.size = 15;
      });
      await context.sync();
    });
  };

  changeFontSize = async () => {
    return await Word.run(async (context) => {
      const range = context.document.getSelection();
      const num = this.state.fontSize || Math.random() * (15 - 11) + 11;
      range.font.size = num;
      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Click ME</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            click me
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.getTooXml}
          >
            get to xml
          </DefaultButton>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.getHtml}>
            get Html
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.contentControl}
          >
            word content control
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.Function}
          >
            extra Functionality
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.getDataFromApi}
          >
            Add Data From Api
          </DefaultButton>
          <input
            type="number"
            placeholder="Enter Font Size"
            onChange={(e) => this.setState({ fontSize: Number(e.target.value) })}
          />
          <button onClick={this.changeFontSize}>Change Font Size</button>
        </HeroList>
      </div>
    );
  }
}

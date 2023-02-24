import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { Configuration, OpenAIApi } from "openai";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

// export interface AppState {
//   listItems: HeroListItem[];
// }

export interface AppState {
  generatedText: string;
  startText: string;
  finalMailText: string;
}

// export default class App extends React.Component<AppProps, AppState> {
//   constructor(props, context) {
//     super(props, context);
//     this.state = {
//       listItems: [],
//     };
//   }

//   componentDidMount() {
//     this.setState({
//       listItems: [
//         {
//           icon: "Ribbon",
//           primaryText: "Achieve more with Office integration",
//         },
//         {
//           icon: "Unlock",
//           primaryText: "Unlock features and functionality",
//         },
//         {
//           icon: "Design",
//           primaryText: "Create and visualize like a pro",
//         },
//       ],
//     });
//   }

//   click = async () => {
//     /**
//      * Insert your Outlook code here
//      */
//   };

//   render() {
//     const { title, isOfficeInitialized } = this.props;

//     if (!isOfficeInitialized) {
//       return (
//         <Progress
//           title={title}
//           logo={require("./../../../assets/logo-filled.png")}
//           message="Please sideload your addin to see app body."
//         />
//       );
//     }

//     return (
//       <div className="ms-welcome">
//         <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
//         <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
//           <p className="ms-font-l">
//             Modify the source files, then click <b>Run</b>.
//           </p>
//           <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
//             Run
//           </DefaultButton>
//         </HeroList>
//       </div>
//     );
//   }
// }

export default class App extends React.Component<AppProps, AppState> {
  constructor(props) {
    super(props);
    this.state = {
      generatedText: "",
      startText: "",
      finalMailText:""
    };
  }
  generateText = async () => {
    var current = this;
    const configuration = new Configuration({
      apiKey: "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",
    });
    const openai = new OpenAIApi(configuration);
    const response = await openai.createCompletion({
      model: "text-davinci-003",
      prompt: "Turn the following text into a professional business mail: " + this.state.startText,
      temperature: 0.7,
      max_tokens: 300,
    });
    current.setState({ generatedText: response.data.choices[0].text });
  };
  insertIntoMail = () => {
    const finalText = this.state.finalMailText.length === 0 ? this.state.generatedText : this.state.finalMailText;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Html,
    });
  };

  render() {
    return (
      <div>
        <main>
          <h2> Open AI business e-mail generator </h2>
          <p>Briefly describe what you want to communicate in the mail:</p>
          <textarea
            onChange={(e) => this.setState({ startText: e.target.value })}
            rows={10}
            cols={40}
          />
          <p>
            <DefaultButton onClick={this.generateText}>
              Generate text
            </DefaultButton>
          </p>
          <textarea
          defaultValue={this.state.generatedText}
          onChange={(e) => this.setState({ finalMailText: e.target.value })}
          rows={10}
          cols={40}
          />
          <p>
            <DefaultButton onClick={this.insertIntoMail}>
              Insert into mail
            </DefaultButton>
          </p>
        </main>
      </div>
    )
  }
}

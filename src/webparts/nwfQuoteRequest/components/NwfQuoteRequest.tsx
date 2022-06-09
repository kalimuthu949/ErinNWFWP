import * as React from "react";
import styles from "./NwfQuoteRequest.module.scss";
import { INwfQuoteRequestProps } from "./INwfQuoteRequestProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { Testname } from "./NWFquotereqnew";
import { NWFQuoteform } from "./NWFQuoteform";
import { sp } from "@pnp/sp/presets/all";

export default class NwfQuoteRequest extends React.Component<
  INwfQuoteRequestProps,
  {}
> {
  constructor(prop: INwfQuoteRequestProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<INwfQuoteRequestProps> {
    return (
      <div>
        <NWFQuoteform
          description={"test"}
          context={this.props.context}
          spcontext={sp.web}
        />
      </div>
    );
  }
}

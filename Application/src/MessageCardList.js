import React, { Component } from "react";
import MessageCard from "./MessageCard";

class MessageCardList extends Component {
  render() {
    return (
      <ul>
        {this.props.messages.map(message => {
          return <MessageCard key={message.id} message={message}></MessageCard>;
        })}
      </ul>
    );
  }
}

export default MessageCardList;

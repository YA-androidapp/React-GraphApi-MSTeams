import React, { Component } from "react";
import AttachmentCard from "./AttachmentCard";

class MessageCard extends Component {
  render() {
    return (
      <li>
        {new Date(
          Date.parse(this.props.message.createdDateTime)
        ).toLocaleDateString()}{" "}
        {JSON.stringify(this.props.message)}
        {this.props.message.id}{" "}
        {this.props.message.replyToId ? this.props.message.replyToId : ""}{" "}
        {this.props.message.etag} {this.props.message.messageType}{" "}
        {new Date(
          Date.parse(this.props.message.createdDateTime)
        ).toLocaleDateString()}
        {this.props.message.lastModifiedDateTime
          ? new Date(
              Date.parse(this.props.message.lastModifiedDateTime)
            ).toLocaleDateString()
          : ""}
        {this.props.message.deletedDateTime
          ? new Date(
              Date.parse(this.props.message.deletedDateTime)
            ).toLocaleDateString()
          : ""}
        {this.props.message.subject ? this.props.message.subject : ""}{" "}
        {this.props.message.summary ? this.props.message.summary : ""}{" "}
        {this.props.message.importance} {this.props.message.locale}{" "}
        {this.props.message.webUrl}{" "}
        {this.props.message.policyViolation
          ? this.props.message.policyViolation
          : ""}{" "}
        {(() => {
          if (this.props.message.from) {
            if (this.props.message.from.application) {
              return <div>{this.props.message.from.application} </div>;
            }
            if (this.props.message.from.device) {
              return <div>{this.props.message.from.device} </div>;
            }
            if (this.props.message.from.conversation) {
              return <div>{this.props.message.from.conversation} </div>;
            }
            if (this.props.message.from.user) {
              return (
                <div>
                  {this.props.message.from.user.id}{" "}
                  {this.props.message.from.user.displayName}{" "}
                  {this.props.message.from.user.userIdentityType}{" "}
                </div>
              );
            }
          }
        })()}
        {(() => {
          if (this.props.message.body) {
            const content = this.props.message.body.content;
            if (this.props.message.body.contentType === "html") {
              return (
                <div>
                  {this.props.message.body.contentType}{" "}
                  <div dangerouslySetInnerHTML={{ __html: content }} />{" "}
                </div>
              );
            } else {
              return (
                <div>
                  {this.props.message.body.contentType} {content}{" "}
                </div>
              );
            }
          }
        })()}
        <ul>
          {(() => {
            if (this.props.message.attachments) {
              this.props.message.attachments.map(attachment => {
                return (
                  <AttachmentCard attachment={attachment}></AttachmentCard>
                );
              });
            }
          })()}
        </ul>
      </li>
    );
  }
}

export default MessageCard;

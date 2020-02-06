import React, { Component } from 'react'

class MessageCard extends Component {
    render() {
        return (
            <li key={this.props.message.id}>
                {new Date(Date.parse(this.props.message.createdDateTime)).toLocaleDateString()}
                {" "}
                {JSON.stringify(this.props.message)}
            </li>
        );
    }
}

export default MessageCard
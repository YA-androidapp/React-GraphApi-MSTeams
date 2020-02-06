import React, { Component } from 'react'
import PropTypes from 'prop-types';

class MessageCardList extends Component {
    render() {
        return (
            <ul>
                {
                    this.props.messages.map(message => {
                        return <li key={message.id}>{message.id} {message.body.content}</li>
                    })
                }
            </ul>
        );
    }
}
MessageCardList.propTypes = {
    onClick: PropTypes.func
}

export default MessageCardList
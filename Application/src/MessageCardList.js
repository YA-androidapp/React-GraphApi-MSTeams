import React, { Component } from 'react'
import PropTypes from 'prop-types';

class MessageCardList extends Component {
    render() {
        return (
            <ul>
            </ul>
        );
    }
}
MessageCardList.propTypes = {
    onClick: PropTypes.func
}

export default MessageCardList
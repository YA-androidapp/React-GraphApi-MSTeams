import React, { Component } from 'react'

class AttachmentCard extends Component {
    render() {
        return (
            <li key={this.props.attachment.id}>
                {' '}
                {JSON.stringify(this.props.attachment)}

                {this.props.attachment.id}{' '}
                {this.props.attachment.contentType}{' '}
                {this.props.attachment.contentUrl}{' '}
                {this.props.attachment.content ? this.props.attachment.content : ''}{' '}
                {this.props.attachment.name}{' '}
                {this.props.attachment.thumbnailUrl ? this.props.attachment.thumbnailUrl : ''}{' '}
            </li>
        );
    }
}

export default AttachmentCard
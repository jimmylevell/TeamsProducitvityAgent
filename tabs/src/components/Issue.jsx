import React, { useState } from "react";
import { Avatar, Card, CardBody, CardHeader, Flex, Text, Button, Dialog } from "@fluentui/react-northstar";
import ReactHtmlParser from 'react-html-parser';
import { useEffect } from "react";

export default function Issue(props) {
  const [comments, setComments] = useState([])

  useEffect(() => {
    const getComments = async () => {
      const URL = "https://dev.azure.com/levellch/kanban-cleanup/_apis/wit/workItems/" + props.item.id + "/comments?api-version=7.0-preview.3"
      let comments = await props.load(URL, "GET", props.token)
      if (comments) {
        comments = comments.comments
        comments = comments.sort((a, b) => new Date(b.modifiedDate) - new Date(a.modifiedDate))
        comments = comments.slice(0, 5)
        setComments(comments)
      }
    }

    getComments()
  }, [props.item.id])

  console.log(comments)

  return (
    <Card fluid style={{ height: "unset" }}>
      <CardHeader>
        <Avatar
          image={props.item.fields["System.AssignedTo"]?.imageUrl}
          name={props.item.fields["System.AssignedTo"]?.displayName}
        />
        <Flex gap="gap.small">
          <Flex column>
            <Text content={props.item.fields["System.Title"]} weight="bold" />
          </Flex>
        </Flex>
      </CardHeader>
      <CardBody>
        <Text content={ReactHtmlParser(props.item.fields["System.Description"])} />

        {comments && comments.length > 0 && (
          <div>
            <h3>Comments</h3>
            {comments.map((comment) => (
              <div key={comment.id}>
                <h4>{comment.modifiedBy.displayName}</h4>
                <p>{comment.text}</p>
              </div>
            ))}
          </div>
        )}
      </CardBody>
      <Card.Footer fitted>
        <Flex space="between">
          <Dialog
            cancelButton="Cancel"
            confirmButton="Save Changes"
            content={ReactHtmlParser(props.item.fields["System.Description"])}
            header={props.item.fields["System.Title"]}
            trigger={<Button content="Show" />}
          />
        </Flex>
      </Card.Footer>
    </Card>
  )
}

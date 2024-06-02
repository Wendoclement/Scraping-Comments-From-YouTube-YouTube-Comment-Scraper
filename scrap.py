import csv
import sys
import datetime
import docx
from googleapiclient.discovery import build


def get_comments(video_id):
    # ^Build the service object
    youtube = build('youtube', 'v3',
                    developerKey='AIzaSyBYLkwuFeKfNgOmF9lj9xICJ2Qqrx666ps')

    # ^List to store the comments, usernames, and replies
    comments_data = []

    # ^Make an initial request
    request = youtube.commentThreads().list(
        part='snippet,replies',
        videoId=video_id,
        maxResults=100,  # You can adjust this value
        textFormat='plainText'
    )

    while request:
        response = request.execute()

        for item in response['items']:
            comment = item['snippet']['topLevelComment']['snippet']['textDisplay']
            author = item['snippet']['topLevelComment']['snippet']['authorDisplayName']
            published_at = item['snippet']['topLevelComment']['snippet']['publishedAt']
            like_count = item['snippet']['topLevelComment']['snippet']['likeCount']
            replies = item['snippet']['totalReplyCount']

            # Convert published_at to a more readable format
            published_at = datetime.datetime.strptime(
                published_at, '%Y-%m-%dT%H:%M:%SZ')
            published_at = published_at.strftime('%Y-%m-%d %H:%M:%S')

            comments_data.append(
                (author, published_at, like_count, replies, comment))

            # ^Check if there are replies to the comment
            if 'replies' in item:
                for reply in item['replies']['comments']:
                    reply_author = reply['snippet']['authorDisplayName']
                    reply_published_at = reply['snippet']['publishedAt']
                    reply_like_count = reply['snippet']['likeCount']

                    # ^Convert reply_published_at to a more readable format
                    reply_published_at = datetime.datetime.strptime(
                        reply_published_at, '%Y-%m-%dT%H:%M:%SZ')
                    reply_published_at = reply_published_at.strftime(
                        '%Y-%m-%d %H:%M:%S')

                    comments_data.append(
                        (reply_author, reply_published_at, reply_like_count, '', reply['snippet']['textDisplay']))

        # ^Check if there are more pages
        if 'nextPageToken' in response:
            request = youtube.commentThreads().list_next(request, response)
        else:
            request = None

# ^New Word Document
    doc = docx.Document()

    # ^Adding a table
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'

    # ^The Header
    header_cells = table.rows[0].cells
    header_cells[0].text = 'No.'
    header_cells[1].text = 'Username'
    header_cells[2].text = 'Published At'
    header_cells[3].text = 'Likes'
    header_cells[4].text = 'Replies'
    header_cells[5].text = 'Comment'

    # ^Inputing the Data in the Tables
    for i, comment in enumerate(comments_data, start=1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(i)
        row_cells[1].text = comment[0]
        row_cells[2].text = comment[1]
        row_cells[3].text = str(comment[2])
        row_cells[4].text = str(comment[3])
        row_cells[5].text = comment[4]

    # ^Saves the document
    doc.save('comments.docx')


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <video_id>")
        sys.exit(1)

    video_id = sys.argv[1]  # ^First command-line argument
    get_comments(video_id)

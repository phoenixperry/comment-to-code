/**
 * Drive comments (Google Docs, and .docx files commented inside Drive).
 * Requires the Drive advanced service, v3.
 */

var COMMENT_FIELDS = 'nextPageToken,comments(id,content,author/displayName,createdTime,' +
  'modifiedTime,resolved,deleted,quotedFileContent/value,' +
  'replies(content,author/displayName,createdTime,deleted))';

/**
 * @param {{id:string,name:string,url:string,sourceType:string}} file
 * @param {Date=} since only comments modified after this (replies count as a modification)
 * @return {Array<Object>} normalised comments
 */
function fetchDriveComments(file, since) {
  var out = [];
  var pageToken;
  do {
    var args = { fields: COMMENT_FIELDS, includeDeleted: true, pageSize: 100 };
    if (since) args.startModifiedTime = since.toISOString();
    if (pageToken) args.pageToken = pageToken;
    var res = Drive.Comments.list(file.id, args);
    (res.comments || []).forEach(function (c) {
      out.push({
        key: c.id,
        fileId: file.id,
        fileName: file.name,
        link: withParam_(file.url, 'disco', c.id),
        excerpt: c.quotedFileContent && c.quotedFileContent.value
          ? decodeEntities(c.quotedFileContent.value) : '[non-text anchor]',
        content: c.content || '',
        author: c.author ? c.author.displayName : '',
        created: c.createdTime,
        modified: c.modifiedTime,
        resolved: !!c.resolved,
        deleted: !!c.deleted,
        replies: (c.replies || []).map(function (r) {
          return {
            content: r.content,
            author: r.author ? r.author.displayName : '',
            created: r.createdTime,
            deleted: !!r.deleted
          };
        }),
        sourceType: file.sourceType
      });
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  return out;
}

function withParam_(url, key, value) {
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + key + '=' + encodeURIComponent(value);
}

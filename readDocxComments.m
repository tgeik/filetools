function comments = readDocxComments(filename)

% comments = readDocxComments(filename)
%
% Extract all comments from a docx file and return in a struct.
%
% INPUT
%   filename    full name of the docx file
%
% OUTPUT
%   comments    struct containing all the comments
%
% Reference
% [1] https://stackoverflow.com/questions/37456361/read-the-data-from-txt-file-inside-zip-file-without-extracting-the-contents-in-m

% create and inputStream from the comments data in the fild
file  = java.io.File(filename);
zipFile = org.apache.tools.zip.ZipFile(file);
commentsEntry = zipFile.getEntry('word/comments.xml');
inputStream = zipFile.getInputStream(commentsEntry);

% read the comments data into variable xmlString
buffer = java.io.ByteArrayOutputStream();
org.apache.commons.io.IOUtils.copy(inputStream, buffer);
xmlString = char(typecast(buffer.toByteArray(), 'uint8')');
inputStream.close;
zipFile.close

% store the XML data in a temp file. readstruct works only on files.
tempFile = fullfile(pwd,'temp.xml');
fid = fopen(tempFile,'Wt');
fwrite(fid,xmlString);
fclose(fid);
% read the temp file into a struct and delete temp file
data = readstruct(tempFile);
delete(tempFile);

% extract only the needed data into a clean struct
c = data.w_comment;
for iComment = numel(c):-1:1
    comments(iComment).author = c(iComment).w_authorAttribute;
    comments(iComment).authorInitials = c(iComment).w_initialsAttribute;
    comments(iComment).date = c(iComment).w_dateAttribute;
    temp = [c(iComment).w_p.w_r.w_t];
    comments(iComment).text = strjoin(temp(2:end),newline);
end

// Daily auto-approval of pending Drive access ("share") requests.
//
// Script properties:
//   driveApprovalFolderIds  - JSON array of folder IDs to auto-approve, oldest to newest.
//                             The last ID is treated as the "latest" duplicate folder.
//   adminEmail              - (optional) email to notify when a request can't be
//                             handled automatically (e.g. the latest folder is also full).
//
// Set up a daily time-driven trigger for autoApproveDriveShareRequests() from the
// Apps Script editor (Triggers > Add trigger > Time-driven > Day timer).

function autoApproveDriveShareRequests() {
  const folderIds = getDriveApprovalFolderIds();
  if (!folderIds.length) {
    Logger.log('autoApproveDriveShareRequests: no folder IDs configured (driveApprovalFolderIds)');
    return;
  }

  const latestFolderId = folderIds[folderIds.length - 1];

  folderIds.forEach((folderId) => {
    let proposals;
    try {
      proposals = Drive.Accessproposals.list(folderId).accessProposals || [];
    } catch (err) {
      Logger.log(`autoApproveDriveShareRequests: failed to list proposals for ${folderId}: ${err.message}`);
      return;
    }

    proposals.forEach((proposal) => resolveDriveAccessProposal(proposal, folderId, latestFolderId));
  });
}

function resolveDriveAccessProposal(proposal, folderId, latestFolderId) {
  const requesterEmail = proposal.requesterEmailAddress;

  try {
    Drive.Accessproposals.resolve({ action: 'ACCEPT', role: ['reader'], sendNotification: false }, folderId, proposal.proposalId);
    Logger.log(`Approved ${requesterEmail} as viewer on ${folderId}`);
    sendApprovalEmail(requesterEmail);
  } catch (err) {
    if (!isCollaboratorLimitError(err)) {
      Logger.log(`autoApproveDriveShareRequests: unexpected error approving ${requesterEmail} on ${folderId}: ${err.message}`);
      notifyDriveApprovalAdmin(`Could not approve ${requesterEmail}'s request on folder ${folderId}:<br>${err.message}`);
      return;
    }

    redirectToLatestFolder(proposal, folderId, latestFolderId);
  }
}

function redirectToLatestFolder(proposal, folderId, latestFolderId) {
  const requesterEmail = proposal.requesterEmailAddress;

  if (folderId === latestFolderId) {
    Logger.log(`autoApproveDriveShareRequests: ${folderId} is already the latest folder and is full; leaving ${requesterEmail}'s request pending`);
    notifyDriveApprovalAdmin(`${requesterEmail} requested access to ${folderId}, which is full and has no newer duplicate to redirect to. Their request was left pending for manual review.`);
    return;
  }

  try {
    Drive.Permissions.create({ role: 'reader', type: 'user', emailAddress: requesterEmail }, latestFolderId, { sendNotificationEmail: false });
  } catch (err) {
    Logger.log(`autoApproveDriveShareRequests: failed to add ${requesterEmail} to latest folder ${latestFolderId}: ${err.message}`);
    notifyDriveApprovalAdmin(`${requesterEmail} requested access to full folder ${folderId}. Tried to add them to the latest duplicate ${latestFolderId} instead, but that failed too:<br>${err.message}`);
    return;
  }

  try {
    Drive.Accessproposals.resolve({ action: 'DENY', sendNotification: false }, folderId, proposal.proposalId);
  } catch (err) {
    Logger.log(`autoApproveDriveShareRequests: added ${requesterEmail} to ${latestFolderId} but failed to decline original request on ${folderId}: ${err.message}`);
  }

  sendDuplicateFolderEmail(requesterEmail, latestFolderId);
  Logger.log(`${requesterEmail} redirected from full folder ${folderId} to latest folder ${latestFolderId}`);
}

function sendDuplicateFolderEmail(requesterEmail, latestFolderId) {
  const link = `https://drive.google.com/drive/folders/${latestFolderId}`;

  MailApp.sendEmail({
    to: requesterEmail,
    subject: 'Your Drive access request',
    htmlBody:
      'Hi,<br><br>The folder you requested access to has reached its maximum number of collaborators, ' +
      'so we\'ve added you to a duplicate version of it instead.<br><br>' +
      `<a href="${link}">Access the folder here</a><br><br>Thanks!`,
  });
}

function sendApprovalEmail(requesterEmail) {
  MailApp.sendEmail({
    to: requesterEmail,
    subject: 'SAT resources folder access approved',
    htmlBody:
      'Here you go! Please let me know if you have any questions. You may want to check out the ' +
      'answer analysis spreadsheet to track progress. Are you a tutor?<br><br>All the best,<br>Danny',
  });
}

function notifyDriveApprovalAdmin(message) {
  const adminEmail = PropertiesService.getScriptProperties().getProperty('adminEmail');
  if (!adminEmail) return;

  MailApp.sendEmail({
    to: adminEmail,
    subject: 'Drive share request needs attention',
    htmlBody: message,
  });
}

function isCollaboratorLimitError(err) {
  const message = ((err && err.message) || '').toLowerCase();
  return message.includes('limit') || message.includes('maximum number');
}

function getDriveApprovalFolderIds() {
  const raw = PropertiesService.getScriptProperties().getProperty('driveApprovalFolderIds');
  return raw ? JSON.parse(raw) : [];
}

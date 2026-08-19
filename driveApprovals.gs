// Daily auto-approval of pending Drive access ("share") requests.
//
// Script properties:
//   driveApprovalFolderIds      - JSON array of folder IDs to auto-approve, oldest to newest.
//                                 The last ID is treated as the "latest" duplicate folder.
//                                 New duplicates are appended here automatically.
//   driveApprovalSourceFolderId - (optional) original folder whose contents new duplicates
//                                 link to. Defaults to DRIVE_APPROVAL_SOURCE_FOLDER_ID below.
//   driveApprovalParentFolderId - (optional) folder to create new duplicates in.
//                                 Defaults to DRIVE_APPROVAL_PARENT_FOLDER_ID below.
//   adminEmail                  - (optional) email to notify when a request can't be
//                                 handled automatically.
//
// Set up a daily time-driven trigger for autoApproveDriveShareRequests() from the
// Apps Script editor (Triggers > Add trigger > Time-driven > Day timer).

const DRIVE_APPROVAL_SOURCE_FOLDER_ID = '1A2POcp4ZdQJroOiZT1pCLaVhgjo8Wm6I';
const DRIVE_APPROVAL_PARENT_FOLDER_ID = '1tHGGSB3Echzw8lnjXh5poWJ0cazQWB_1';

function autoApproveDriveShareRequests() {
  const folderIds = getDriveApprovalFolderIds();
  if (!folderIds.length) {
    Logger.log('autoApproveDriveShareRequests: no folder IDs configured (driveApprovalFolderIds)');
    return;
  }

  let latestFolderId = folderIds[folderIds.length - 1];

  folderIds.forEach((folderId) => {
    let proposals;
    try {
      proposals = Drive.Accessproposals.list(folderId).accessProposals || [];
    } catch (err) {
      Logger.log(`autoApproveDriveShareRequests: failed to list proposals for ${folderId}: ${err.message}`);
      return;
    }

    // A duplicate created part way through this run becomes the target for the
    // requests still to be processed.
    proposals.forEach((proposal) => {
      latestFolderId = resolveDriveAccessProposal(proposal, folderId, latestFolderId);
    });
  });
}

function resolveDriveAccessProposal(proposal, folderId, latestFolderId) {
  const requesterEmail = proposal.requesterEmailAddress;

  try {
    Drive.Accessproposals.resolve({ action: 'ACCEPT', role: ['reader'], sendNotification: false }, folderId, proposal.proposalId);
    Logger.log(`Approved ${requesterEmail} as viewer on ${folderId}`);
    sendApprovalEmail(requesterEmail);
    return latestFolderId;
  } catch (err) {
    if (!isCollaboratorLimitError(err)) {
      Logger.log(`autoApproveDriveShareRequests: unexpected error approving ${requesterEmail} on ${folderId}: ${err.message}`);
      notifyDriveApprovalAdmin(`Could not approve ${requesterEmail}'s request on folder ${folderId}:<br>${err.message}`);
      return latestFolderId;
    }

    return redirectToLatestFolder(proposal, folderId, latestFolderId);
  }
}

// Moves a request off a full folder and onto the newest duplicate, creating that
// duplicate first if there isn't one to move to. Returns the latest folder ID,
// which changes whenever a duplicate is created.
function redirectToLatestFolder(proposal, folderId, latestFolderId) {
  const requesterEmail = proposal.requesterEmailAddress;
  let targetFolderId = latestFolderId;

  if (targetFolderId === folderId) {
    targetFolderId = createDuplicateFolder();
    if (!targetFolderId) {
      Logger.log(`redirectToLatestFolder: ${folderId} is full and creating a duplicate failed; leaving ${requesterEmail}'s request pending`);
      notifyDriveApprovalAdmin(`${requesterEmail} requested access to ${folderId}, which is full, and creating a new duplicate folder failed. Their request was left pending for manual review.`);
      return latestFolderId;
    }
  }

  try {
    addReaderToFolder(requesterEmail, targetFolderId);
  } catch (err) {
    if (!isCollaboratorLimitError(err)) {
      Logger.log(`redirectToLatestFolder: failed to add ${requesterEmail} to ${targetFolderId}: ${err.message}`);
      notifyDriveApprovalAdmin(`${requesterEmail} requested access to full folder ${folderId}. Tried to add them to the latest duplicate ${targetFolderId} instead, but that failed too:<br>${err.message}`);
      return targetFolderId;
    }

    // The latest duplicate is full as well, so start another one and try once more.
    const newFolderId = createDuplicateFolder();
    if (!newFolderId) {
      Logger.log(`redirectToLatestFolder: ${targetFolderId} is also full and creating a duplicate failed; leaving ${requesterEmail}'s request pending`);
      notifyDriveApprovalAdmin(`${requesterEmail} requested access to full folder ${folderId}. The latest duplicate ${targetFolderId} is full too, and creating a new one failed. Their request was left pending for manual review.`);
      return targetFolderId;
    }

    targetFolderId = newFolderId;

    try {
      addReaderToFolder(requesterEmail, targetFolderId);
    } catch (retryErr) {
      Logger.log(`redirectToLatestFolder: failed to add ${requesterEmail} to new folder ${targetFolderId}: ${retryErr.message}`);
      notifyDriveApprovalAdmin(`Created duplicate folder ${targetFolderId} for ${requesterEmail}, but adding them to it failed:<br>${retryErr.message}`);
      return targetFolderId;
    }
  }

  try {
    Drive.Accessproposals.resolve({ action: 'DENY', sendNotification: false }, folderId, proposal.proposalId);
  } catch (err) {
    Logger.log(`redirectToLatestFolder: added ${requesterEmail} to ${targetFolderId} but failed to decline original request on ${folderId}: ${err.message}`);
  }

  sendDuplicateFolderEmail(requesterEmail, targetFolderId);
  Logger.log(`${requesterEmail} redirected from full folder ${folderId} to latest folder ${targetFolderId}`);
  return targetFolderId;
}

// Creates the next duplicate of the resources folder and registers it as the latest.
// Returns the new folder ID, or null if it couldn't be created.
function createDuplicateFolder() {
  const sourceFolderId = getDriveApprovalSourceFolderId();

  try {
    const source = Drive.Files.get(sourceFolderId, { fields: 'id, name' });
    const version = getDriveApprovalFolderIds().length + 1;

    const newFolder = Drive.Files.create({
      name: `${source.name} (${version})`,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [getDriveApprovalParentFolderId()],
    });
    copyFolderAsShortcuts(sourceFolderId, newFolder.id);
    appendDriveApprovalFolderId(newFolder.id);

    Logger.log(`Created duplicate resources folder ${newFolder.name} (${newFolder.id})`);
    notifyDriveApprovalAdmin(
      `The previous resources folder filled up, so a new duplicate was created: ` +
      `<a href="https://drive.google.com/drive/folders/${newFolder.id}">${newFolder.name}</a>. ` +
      `It was added to driveApprovalFolderIds and new requests are being sent there.`
    );

    return newFolder.id;
  } catch (err) {
    Logger.log(`createDuplicateFolder: failed to duplicate ${sourceFolderId}: ${err.message}`);
    return null;
  }
}

function copyFolderAsShortcuts(sourceFolderId, destinationFolderId) {
  let pageToken = null;

  do {
    const response = Drive.Files.list({
      q: `'${sourceFolderId}' in parents and trashed = false`,
      fields: 'nextPageToken, files(id, name, mimeType, shortcutDetails)',
      pageSize: 100,
      pageToken: pageToken,
    });

    (response.files || []).forEach((file) => {
      // Drive won't point a shortcut at another shortcut, so if the original folder
      // ever holds one, follow it through to the item it targets.
      const targetId = file.shortcutDetails ? file.shortcutDetails.targetId : file.id;

      Drive.Files.create({
        name: file.name,
        mimeType: 'application/vnd.google-apps.shortcut',
        parents: [destinationFolderId],
        shortcutDetails: { targetId: targetId },
      });
    });

    pageToken = response.nextPageToken;
  } while (pageToken);
}

function addReaderToFolder(requesterEmail, folderId) {
  Drive.Permissions.create({ role: 'reader', type: 'user', emailAddress: requesterEmail }, folderId, { sendNotificationEmail: false });
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

function getDriveApprovalSourceFolderId() {
  return PropertiesService.getScriptProperties().getProperty('driveApprovalSourceFolderId') || DRIVE_APPROVAL_SOURCE_FOLDER_ID;
}

function getDriveApprovalParentFolderId() {
  return PropertiesService.getScriptProperties().getProperty('driveApprovalParentFolderId') || DRIVE_APPROVAL_PARENT_FOLDER_ID;
}

function appendDriveApprovalFolderId(folderId) {
  const folderIds = getDriveApprovalFolderIds();
  folderIds.push(folderId);
  PropertiesService.getScriptProperties().setProperty('driveApprovalFolderIds', JSON.stringify(folderIds));
}

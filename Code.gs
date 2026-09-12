function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Chat');
}

/* ========= UTIL ========= */

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function normalizeUsername(value) {
  return String(value || '').trim();
}

function hashPassword(password) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password
  );
  return Utilities.base64Encode(raw);
}

function getSessionCache() {
  return CacheService.getScriptCache();
}

function createSession(username) {
  const token = Utilities.getUuid();
  getSessionCache().put('chat_session_' + token, normalizeUsername(username), 1800);
  return token;
}

function getUserFromSession(token) {
  if (!token) return null;
  return normalizeUsername(getSessionCache().get('chat_session_' + String(token)) || '');
}

function clearSession(token) {
  if (token) {
    getSessionCache().remove('chat_session_' + String(token));
  }
}

function requireSessionUser(sessionToken) {
  const username = getUserFromSession(sessionToken);
  if (!username) {
    throw new Error('Session expired or invalid');
  }
  return username;
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function imageDataToBlob(imageData) {
  if (!imageData || typeof imageData !== 'string') return null;

  const matches = imageData.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!matches) return null;

  const mimeType = matches[1];
  const base64 = matches[2];
  const extension = mimeType.includes('png') ? 'png' : mimeType.includes('gif') ? 'gif' : 'jpg';
  const fileName = Utilities.getUuid() + '.' + extension;
  return Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
}

function getRoomMessagesFromSheet(roomId, minCreatedAt) {
  const room = String(roomId || '').trim();
  if (!room) return [];

  const minTime = Number(minCreatedAt) || 0;
  const rows = sheet_('Messages', ['Id', 'RoomId', 'Username', 'Message', 'Image', 'CreatedAt', 'MessageType'])
    .getDataRange()
    .getValues()
    .slice(1);

  return rows
    .filter(row => String(row[1]) === room && Number(row[5] || 0) > minTime)
    .map(row => ({
      id: String(row[0]),
      roomId: String(row[1]),
      username: String(row[2]),
      text: String(row[3] || ''),
      image: String(row[4] || ''),
      createdAt: Number(row[5] || 0),
      type: String(row[6] || (row[4] ? 'image' : 'text'))
    }))
    .sort((a, b) => a.createdAt - b.createdAt);
}

function getImageBase64Size(imageData) {
  if (!imageData || typeof imageData !== 'string') return 0;
  const match = imageData.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  return match ? match[2].length : String(imageData).replace(/\s+/g, '').length;
}

function normalizeImageBase64(imageData) {
  if (!imageData || typeof imageData !== 'string') return '';
  const match = imageData.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  return match ? match[2] : String(imageData).replace(/\s+/g, '');
}

function isRoomMember(roomId, username) {
  if (!roomId || !username) return false;

  const memberSheet = sheet_('RoomMembers', ['RoomId', 'Username', 'JoinedAt']);
  const rows = memberSheet.getDataRange().getValues().slice(1);

  return rows.some(row =>
    String(row[0]) === String(roomId) &&
    String(row[1]).toLowerCase() === String(username).toLowerCase()
  );
}

function ensureRoomMember(roomId, username) {
  if (!roomId || !username) return;

  const memberSheet = sheet_('RoomMembers', ['RoomId', 'Username', 'JoinedAt']);
  const rows = memberSheet.getDataRange().getValues().slice(1);
  const alreadyMember = rows.some(row =>
    String(row[0]) === String(roomId) && String(row[1]).toLowerCase() === String(username).toLowerCase()
  );

  if (!alreadyMember) {
    memberSheet.appendRow([roomId, username, Date.now()]);
  }
}

function getRoomById(roomId) {
  const rooms = sheet_('Rooms', ['RoomId', 'RoomName', 'Owner', 'IsPrivate', 'CreatedAt', 'RoomType'])
    .getDataRange()
    .getValues()
    .slice(1);

  return rooms.find(row => String(row[0]) === String(roomId)) || null;
}

/* ========= AUTH ========= */

function validateSession(sessionToken) {
  try {
    const username = requireSessionUser(sessionToken);
    return { ok: true, username: username };
  } catch (err) {
    return { ok: false };
  }
}

function login(username, password) {
  const user = normalizeUsername(username);
  if (!user || !password) return { ok: false };

  const sh = sheet_('Users', ['Username', 'Password', 'Active', 'CreatedAt', 'Visible']);
  const rows = sh.getDataRange().getValues().slice(1);
  const hash = hashPassword(password);

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === user.toLowerCase()) {
      if (String(rows[i][1]) !== hash) return { ok: false };
      sh.getRange(i + 2, 3).setValue(true);
      if (rows[i].length < 5) {
        sh.getRange(i + 2, 5).setValue(true);
      }
      return { ok: true, token: createSession(user), username: user };
    }
  }

  return { ok: false };
}

function registerAccount(username, password) {
  const user = normalizeUsername(username);
  if (!user || !password) return { ok: false, error: 'Missing username or password.' };

  const sh = sheet_('Users', ['Username', 'Password', 'Active', 'CreatedAt', 'Visible']);
  const rows = sh.getDataRange().getValues().slice(1);
  const existing = rows.some(row => String(row[0]).toLowerCase() === user.toLowerCase());

  if (existing) {
    return { ok: false, error: 'That username already exists.' };
  }

  sh.appendRow([user, hashPassword(password), true, Date.now(), true]);
  return { ok: true, username: user };
}

function logout(sessionToken) {
  clearSession(sessionToken);
  return { ok: true };
}

function getUsers(sessionToken) {
  requireSessionUser(sessionToken);
  const sh = sheet_('Users', ['Username', 'Password', 'Active', 'CreatedAt', 'Visible']);
  const rows = sh.getDataRange().getValues().slice(1);

  return rows
    .filter(row => String(row[0] || '').trim() && String(row[4] || 'true').toLowerCase() !== 'false')
    .map(row => ({ username: String(row[0]) }));
}

function getUserVisibility(sessionToken) {
  const username = requireSessionUser(sessionToken);
  const sh = sheet_('Users', ['Username', 'Password', 'Active', 'CreatedAt', 'Visible']);
  const rows = sh.getDataRange().getValues().slice(1);
  const row = rows.find(r => String(r[0]).toLowerCase() === username.toLowerCase());
  if (!row) return { visible: true };
  return { visible: String(row[4] || 'true').toLowerCase() !== 'false' };
}

function updateUserVisibility(visible, sessionToken) {
  const username = requireSessionUser(sessionToken);
  const sh = sheet_('Users', ['Username', 'Password', 'Active', 'CreatedAt', 'Visible']);
  const rows = sh.getDataRange().getValues().slice(1);
  const index = rows.findIndex(r => String(r[0]).toLowerCase() === username.toLowerCase());

  if (index === -1) {
    return { ok: false };
  }

  sh.getRange(index + 2, 5).setValue(!!visible);
  return { ok: true, visible: !!visible };
}

/* ========= ROOMS ========= */

function createRoom(name, isPrivate, invites, sessionToken) {
  const roomOwner = requireSessionUser(sessionToken);
  const roomName = String(name || '').trim() || 'New room';
  const roomId = Utilities.getUuid();

  const roomSheet = sheet_('Rooms', ['RoomId', 'RoomName', 'Owner', 'IsPrivate', 'CreatedAt', 'RoomType']);
  roomSheet.appendRow([roomId, roomName, roomOwner, !!isPrivate, Date.now(), 'group']);
  ensureRoomMember(roomId, roomOwner);

  const inviteList = Array.isArray(invites) ? invites : [];
  inviteList.forEach(invitee => {
    if (invitee) inviteUser(roomId, invitee, sessionToken);
  });

  return { ok: true, roomId: roomId };
}

function getOrCreateDirectRoom(targetUser, sessionToken) {
  const currentUser = requireSessionUser(sessionToken);
  const otherUser = normalizeUsername(targetUser);

  if (!otherUser || otherUser.toLowerCase() === currentUser.toLowerCase()) {
    return null;
  }

  const participants = [currentUser, otherUser].sort((x, y) => x.localeCompare(y));
  const roomId = 'dm_' + participants.join('_');
  const roomSheet = sheet_('Rooms', ['RoomId', 'RoomName', 'Owner', 'IsPrivate', 'CreatedAt', 'RoomType']);
  const existing = getRoomById(roomId);

  if (!existing) {
    roomSheet.appendRow([roomId, participants.join(' / '), currentUser, true, Date.now(), 'dm']);
  }

  ensureRoomMember(roomId, currentUser);
  ensureRoomMember(roomId, otherUser);

  return { id: roomId, name: participants.join(' / ') };
}

function getRooms(sessionToken) {
  const currentUser = requireSessionUser(sessionToken);

  const roomRows = sheet_('Rooms', ['RoomId', 'RoomName', 'Owner', 'IsPrivate', 'CreatedAt', 'RoomType'])
    .getDataRange()
    .getValues()
    .slice(1);

  const memberRows = sheet_('RoomMembers', ['RoomId', 'Username', 'JoinedAt'])
    .getDataRange()
    .getValues()
    .slice(1);

  const membership = new Set(
    memberRows
      .filter(row => String(row[1]).toLowerCase() === currentUser.toLowerCase())
      .map(row => String(row[0]))
  );

  return roomRows
    .filter(row => {
      const roomId = String(row[0]);
      const roomType = String(row[5] || 'group');
      const isPrivate = !!row[3];
      const owner = String(row[2] || '').toLowerCase();

      if (roomType === 'dm') {
        return membership.has(roomId);
      }

      if (!isPrivate) return true;
      return owner === currentUser.toLowerCase() || membership.has(roomId);
    })
    .map(row => {
      const roomId = String(row[0]);
      const roomType = String(row[5] || 'group');
      const roomName = String(row[1] || 'Room');

      if (roomType === 'dm') {
        const other = memberRows
          .filter(member => String(member[0]) === roomId && String(member[1]).toLowerCase() !== currentUser.toLowerCase())
          .map(member => member[1])[0] || roomName;

        return {
          id: roomId,
          name: other,
          type: 'dm',
          private: true,
          owner: String(row[2] || '')
        };
      }

      return {
        id: roomId,
        name: roomName,
        type: 'group',
        private: !!row[3],
        owner: String(row[2] || '')
      };
    });
}

function inviteUser(roomId, username, sessionToken) {
  const currentUser = requireSessionUser(sessionToken);
  const invitedUser = normalizeUsername(username);
  const cleanRoomId = String(roomId || '').trim();

  if (!cleanRoomId || !invitedUser) return false;

  const room = getRoomById(cleanRoomId);
  if (!room) return false;

  if (String(room[2]) !== currentUser && !isRoomMember(cleanRoomId, currentUser)) {
    return false;
  }

  const invites = sheet_('Invites', ['RoomId', 'Username', 'CreatedAt']);
  const existing = invites.getDataRange().getValues().slice(1);
  const alreadyInvited = existing.some(row =>
    String(row[0]) === cleanRoomId && String(row[1]).toLowerCase() === invitedUser.toLowerCase()
  );

  if (!alreadyInvited) {
    invites.appendRow([cleanRoomId, invitedUser, Date.now()]);
  }

  ensureRoomMember(cleanRoomId, invitedUser);
  return true;
}

function deleteRoom(roomId, sessionToken) {
  const currentUser = requireSessionUser(sessionToken);
  const cleanRoomId = String(roomId || '').trim();
  if (!cleanRoomId) return { ok: false };

  const room = getRoomById(cleanRoomId);
  if (!room) return { ok: false };

  const roomType = String(room[5] || 'group');
  const isOwner = String(room[2] || '').toLowerCase() === currentUser.toLowerCase();
  const isDm = roomType === 'dm';

  if (!isOwner && !isDm) return { ok: false };

  const roomSheet = sheet_('Rooms', ['RoomId', 'RoomName', 'Owner', 'IsPrivate', 'CreatedAt', 'RoomType']);
  const roomRows = roomSheet.getDataRange().getValues();
  const keepRows = roomRows.filter(row => String(row[0]) !== cleanRoomId);

  roomSheet.clear();
  if (keepRows.length) {
    roomSheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
  }

  const memberSheet = sheet_('RoomMembers', ['RoomId', 'Username', 'JoinedAt']);
  const memberRows = memberSheet.getDataRange().getValues();
  const keepMembers = memberRows.filter(row => String(row[0]) !== cleanRoomId);
  memberSheet.clear();
  if (keepMembers.length) {
    memberSheet.getRange(1, 1, keepMembers.length, keepMembers[0].length).setValues(keepMembers);
  }

  const messageSheet = sheet_('Messages', ['Id', 'RoomId', 'Username', 'Message', 'Image', 'CreatedAt', 'MessageType']);
  const messageRows = messageSheet.getDataRange().getValues();
  const keepMessages = messageRows.filter(row => String(row[1]) !== cleanRoomId);
  messageSheet.clear();
  if (keepMessages.length) {
    messageSheet.getRange(1, 1, keepMessages.length, keepMessages[0].length).setValues(keepMessages);
  }

  const invitesSheet = sheet_('Invites', ['RoomId', 'Username', 'CreatedAt']);
  const inviteRows = invitesSheet.getDataRange().getValues();
  const keepInvites = inviteRows.filter(row => String(row[0]) !== cleanRoomId);
  invitesSheet.clear();
  if (keepInvites.length) {
    invitesSheet.getRange(1, 1, keepInvites.length, keepInvites[0].length).setValues(keepInvites);
  }

  return { ok: true };
}

/* ========= MESSAGES ========= */

function saveMessage(roomId, text, imageData, sessionToken) {
  const sender = requireSessionUser(sessionToken);
  const room = String(roomId || '').trim();
  const messageText = String(text || '').trim();

  if (!room || !sender || (!messageText && !imageData)) return null;
  if (!isRoomMember(room, sender)) return null;

  const normalizedImage = normalizeImageBase64(imageData);
  const imageBase64 = normalizedImage && getImageBase64Size(normalizedImage) <= 50000 ? normalizedImage : '';
  const messageId = Utilities.getUuid();
  const createdAt = Date.now();

  const savedMessage = {
    id: messageId,
    roomId: room,
    username: sender,
    text: messageText,
    image: imageBase64,
    createdAt: createdAt,
    type: imageBase64 ? 'image' : 'text'
  };

  sheet_('Messages', ['Id', 'RoomId', 'Username', 'Message', 'Image', 'CreatedAt', 'MessageType'])
    .appendRow([
      savedMessage.id,
      savedMessage.roomId,
      savedMessage.username,
      savedMessage.text,
      savedMessage.image,
      savedMessage.createdAt,
      savedMessage.type
    ]);

  return savedMessage;
}

function uploadImage(imageData) {
  return imageData || '';
}

function getMessages(roomId, beforeCreatedAt, limit, sessionToken, sinceCreatedAt) {
  const room = String(roomId || '').trim();
  const currentUser = requireSessionUser(sessionToken);

  if (!room || !isRoomMember(room, currentUser)) {
    return {
      messages: [],
      hasMoreOlder: false
    };
  }

  const pageSize = Math.max(1, Math.min(Number(limit) || 7, 50));

  const messageSheet = sheet_(
    'Messages',
    ['Id', 'RoomId', 'Username', 'Message', 'Image', 'CreatedAt', 'MessageType']
  );

  const lastRow = messageSheet.getLastRow();

  if (lastRow <= 1) {
    return {
      messages: [],
      hasMoreOlder: false
    };
  }

  /*
   * IMPORTANT:
   * For polling, only inspect the newest rows.
   *
   * Messages are appended chronologically, so we can walk backward
   * from the bottom instead of processing the entire sheet.
   */
  if (
    sinceCreatedAt !== undefined &&
    sinceCreatedAt !== null &&
    sinceCreatedAt !== ''
  ) {
    const since = Number(sinceCreatedAt) || 0;

    const firstDataRow = Math.max(2, lastRow - 500 + 1);
    const rowCount = lastRow - firstDataRow + 1;

    const rows = messageSheet
      .getRange(firstDataRow, 1, rowCount, 7)
      .getValues();

    const messages = [];

    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];

      const createdAt = Number(row[5] || 0);

      /*
       * Use >= rather than >.
       *
       * This prevents two messages created during the same
       * millisecond from causing one to disappear.
       */
      if (createdAt < since) {
        break;
      }

      if (String(row[1]) !== room) {
        continue;
      }

      messages.push({
        id: String(row[0]),
        roomId: String(row[1]),
        username: String(row[2]),
        text: String(row[3] || ''),
        image: String(row[4] || ''),
        createdAt: createdAt,
        type: String(row[6] || (row[4] ? 'image' : 'text'))
      });

      if (messages.length >= 50) {
        break;
      }
    }

    messages.reverse();

    return {
      messages: messages,
      hasMoreOlder: false
    };
  }

  /*
   * Initial load / older messages.
   *
   * These requests are less frequent, so we can read the sheet normally.
   */
  const rows = messageSheet
    .getDataRange()
    .getValues()
    .slice(1);

  const messages = rows
    .filter(row => String(row[1]) === room)
    .map(row => ({
      id: String(row[0]),
      roomId: String(row[1]),
      username: String(row[2]),
      text: String(row[3] || ''),
      image: String(row[4] || ''),
      createdAt: Number(row[5] || 0),
      type: String(row[6] || (row[4] ? 'image' : 'text'))
    }))
    .sort((a, b) =>
      Number(a.createdAt || 0) -
      Number(b.createdAt || 0)
    );

  if (
    beforeCreatedAt !== undefined &&
    beforeCreatedAt !== null &&
    beforeCreatedAt !== ''
  ) {
    const beforeValue = Number(beforeCreatedAt) || 0;

    const olderMessages = messages.filter(
      msg => Number(msg.createdAt || 0) < beforeValue
    );

    const page = olderMessages.slice(-pageSize);

    return {
      messages: page,
      hasMoreOlder: olderMessages.length > page.length
    };
  }

  const latestPage = messages.slice(-pageSize);

  return {
    messages: latestPage,
    hasMoreOlder: messages.length > latestPage.length
  };
}

function handleMarkModuleCompleted(params) {
  const moduleID = params.moduleID;
  const trainerID = params.trainerID;
  const notes = params.notes || '';
  
  const sheet = getSheet(SHEETS.MODULES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === moduleID) {
      const row = i + 1;
      const batchID = data[i][1];
      const moduleName = data[i][2];
      const actualDate = getCurrentDate();
      
      sheet.getRange(row, 4).setValue('COMPLETED');
      sheet.getRange(row, 6).setValue(actualDate);
      sheet.getRange(row, 8).setValue(notes);
      
      const studentSheet = getSheet(SHEETS.STUDENTS);
      const studentData = studentSheet.getDataRange().getValues();
      
      for (let j = 1; j < studentData.length; j++) {
        if (studentData[j][4] === batchID && studentData[j][6] === 'Active') {
          createNotification(
            studentData[j][3],
            'STUDENT',
            'Module Completed',
            `Module "${moduleName}" has been completed by your trainer.`,
            batchID
          );
        }
      }
      
      const adminEmail = getConfigValue('AdminEmail');
      if (adminEmail) {
        createNotification(adminEmail, 'ADMIN', 'Module Completed',
          `Module ${moduleName} completed in batch ${batchID} by trainer ${trainerID}`, batchID);
      }
      
      logAudit('MODULE_COMPLETED', 'IN_PROGRESS', 'COMPLETED', batchID, trainerID, '', trainerID, 'MODULE');
      
      return jsonResponse(true, 'Module marked as completed');
    }
  }
  
  return jsonResponse(false, 'Module not found');
}

function handleRescheduleModule(params) {
  const moduleID = params.moduleID;
  const newDate = params.newDate;
  const trainerID = params.trainerID;
  
  const sheet = getSheet(SHEETS.MODULES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === moduleID) {
      const row = i + 1;
      const oldDate = data[i][4];
      const batchID = data[i][1];
      const moduleName = data[i][2];
      
      sheet.getRange(row, 5).setValue(newDate);
      
      const studentSheet = getSheet(SHEETS.STUDENTS);
      const studentData = studentSheet.getDataRange().getValues();
      
      for (let j = 1; j < studentData.length; j++) {
        if (studentData[j][4] === batchID && studentData[j][6] === 'Active') {
          createNotification(
            studentData[j][3],
            'STUDENT',
            'Module Rescheduled',
            `Module "${moduleName}" rescheduled from ${oldDate} to ${newDate}`,
            batchID
          );
        }
      }
      
      logAudit('MODULE_RESCHEDULED', oldDate, newDate, batchID, trainerID, '', trainerID, 'MODULE');
      
      return jsonResponse(true, 'Module rescheduled successfully');
    }
  }
  
  return jsonResponse(false, 'Module not found');
}

// ==================== ADMIN ENDPOINTS ====================

function verifyAdmin(adminKey) {
  return adminKey === ADMIN_SECRET_KEY;
}

function handleAdminAddStudent(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.STUDENTS);
  const studentID = generateID('STD');
  
  sheet.appendRow([
    studentID,
    sanitizeInput(params.name),
    sanitizeInput(params.phone),
    sanitizeInput(params.email),
    sanitizeInput(params.batchID),
    '',
    'Active',
    getCurrentDate()
  ]);
  
  createNotification(params.email, 'STUDENT', 'Welcome to DA360',
    `Welcome ${params.name}! Your Student ID is ${studentID}. Please login to access your portal.`, params.batchID);
  
  logAudit('STUDENT_ADDED', '', studentID, params.batchID, '', '', params.adminEmail, 'STUDENT');
  logChangeHistory(params.adminEmail, 'STUDENT_ADDED', {
    studentID: studentID,
    name: params.name,
    batchID: params.batchID
  });
  
  return jsonResponse(true, 'Student added successfully', { studentID: studentID });
}

function handleAdminUpdateStudent(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.STUDENTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.studentID) {
      const row = i + 1;
      
      if (params.name) sheet.getRange(row, 2).setValue(sanitizeInput(params.name));
      if (params.phone) sheet.getRange(row, 3).setValue(sanitizeInput(params.phone));
      if (params.email) sheet.getRange(row, 4).setValue(sanitizeInput(params.email));
      if (params.status) sheet.getRange(row, 7).setValue(params.status);
      
      logAudit('STUDENT_UPDATED', '', params.studentID, '', '', '', params.adminEmail, 'STUDENT');
      
      return jsonResponse(true, 'Student updated successfully');
    }
  }
  
  return jsonResponse(false, 'Student not found');
}

function handleAdminMoveStudent(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.STUDENTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.studentID) {
      const row = i + 1;
      const oldBatchID = data[i][4];
      const newBatchID = params.newBatchID;
      
      sheet.getRange(row, 5).setValue(newBatchID);
      
      createNotification(data[i][3], 'STUDENT', 'Batch Changed',
        `You have been moved from batch ${oldBatchID} to ${newBatchID}`, newBatchID);
      
      logAudit('STUDENT_MOVED', oldBatchID, newBatchID, newBatchID, '', '', params.adminEmail, 'STUDENT');
      logChangeHistory(params.adminEmail, 'STUDENT_MOVED', {
        studentID: params.studentID,
        oldBatch: oldBatchID,
        newBatch: newBatchID
      });
      
      return jsonResponse(true, 'Student moved successfully');
    }
  }
  
  return jsonResponse(false, 'Student not found');
}

function handleAdminAddTrainer(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.TRAINERS);
  const trainerID = generateID('TR');
  
  sheet.appendRow([
    trainerID,
    sanitizeInput(params.name),
    sanitizeInput(params.email),
    sanitizeInput(params.phone),
    'Active'
  ]);
  
  createNotification(params.email, 'TRAINER', 'Welcome to DA360',
    `Welcome ${params.name}! Your Trainer ID is ${trainerID}. You can now access the trainer portal.`, '');
  
  logAudit('TRAINER_ADDED', '', trainerID, '', trainerID, '', params.adminEmail, 'TRAINER');
  
  return jsonResponse(true, 'Trainer added successfully', { trainerID: trainerID });
}

function handleAdminAddBatch(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.BATCHES);
  const batchID = generateID('BATCH');
  
  sheet.appendRow([
    batchID,
    sanitizeInput(params.name),
    sanitizeInput(params.trainerID),
    sanitizeInput(params.classroomID),
    sanitizeInput(params.mode),
    params.startDate,
    params.endDate,
    '',
    'Active',
    params.notes || ''
  ]);
  
  if (params.timingSlotID) {
    const timingSheet = getSheet(SHEETS.BATCH_TIMINGS);
    const batchTimingID = generateID('BT');
    timingSheet.appendRow([
      batchTimingID,
      batchID,
      params.timingSlotID,
      params.startDate,
      params.endDate,
      'Active'
    ]);
  }
  
  const trainerSheet = getSheet(SHEETS.TRAINERS);
  const trainerData = trainerSheet.getDataRange().getValues();
  for (let i = 1; i < trainerData.length; i++) {
    if (trainerData[i][0] === params.trainerID) {
      createNotification(trainerData[i][2], 'TRAINER', 'New Batch Assigned',
        `You have been assigned to batch ${params.name} (${batchID})`, batchID);
      break;
    }
  }
  
  logAudit('BATCH_CREATED', '', batchID, batchID, params.trainerID, params.classroomID, params.adminEmail, 'BATCH');
  logChangeHistory(params.adminEmail, 'BATCH_CREATED', {
    batchID: batchID,
    name: params.name,
    trainerID: params.trainerID,
    classroomID: params.classroomID,
    mode: params.mode
  });
  
  return jsonResponse(true, 'Batch created successfully', { batchID: batchID });
}

function handleAdminUpdateBatch(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.BATCHES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.batchID) {
      const row = i + 1;
      
      if (params.name) sheet.getRange(row, 2).setValue(sanitizeInput(params.name));
      if (params.startDate) sheet.getRange(row, 6).setValue(params.startDate);
      if (params.endDate) sheet.getRange(row, 7).setValue(params.endDate);
      if (params.status) sheet.getRange(row, 9).setValue(params.status);
      if (params.notes) sheet.getRange(row, 10).setValue(params.notes);
      
      logAudit('BATCH_UPDATED', '', params.batchID, params.batchID, '', '', params.adminEmail, 'BATCH');
      
      return jsonResponse(true, 'Batch updated successfully');
    }
  }
  
  return jsonResponse(false, 'Batch not found');
}

function handleAdminMergeBatch(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sourceBatchIDs = params.sourceBatchIDs;
  const targetBatchID = generateID('BATCH');
  const targetBatchName = params.targetBatchName;
  const trainerID = params.trainerID;
  const classroomID = params.classroomID;
  const mode = params.mode;
  
  const batchSheet = getSheet(SHEETS.BATCHES);
  batchSheet.appendRow([
    targetBatchID,
    targetBatchName,
    trainerID,
    classroomID,
    mode,
    getCurrentDate(),
    params.endDate || '',
    '',
    'Active',
    'Merged batch from: ' + sourceBatchIDs.join(', ')
  ]);
  
  const studentSheet = getSheet(SHEETS.STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  const movedStudents = [];
  
  for (let i = 1; i < studentData.length; i++) {
    if (sourceBatchIDs.includes(studentData[i][4])) {
      const row = i + 1;
      studentSheet.getRange(row, 5).setValue(targetBatchID);
      
      createNotification(studentData[i][3], 'STUDENT', 'Batch Merged',
        `Your batch has been merged into ${targetBatchName}. New batch ID: ${targetBatchID}`, targetBatchID);
      
      movedStudents.push(studentData[i][0]);
    }
  }
  
  const moduleSheet = getSheet(SHEETS.MODULES);
  const moduleData = moduleSheet.getDataRange().getValues();
  
  for (let i = 1; i < moduleData.length; i++) {
    if (sourceBatchIDs.includes(moduleData[i][1]) && moduleData[i][3] !== 'COMPLETED') {
      const newModuleID = generateID('MOD');
      moduleSheet.appendRow([
        newModuleID,
        targetBatchID,
        moduleData[i][2],
        moduleData[i][3],
        moduleData[i][4],
        moduleData[i][5],
        trainerID,
        'Merged from ' + moduleData[i][1]
      ]);
    }
  }
  
  const batchData = batchSheet.getDataRange().getValues();
  for (let i = 1; i < batchData.length; i++) {
    if (sourceBatchIDs.includes(batchData[i][0])) {
      const row = i + 1;
      batchSheet.getRange(row, 9).setValue('Merged');
    }
  }
  
  const trainerSheet = getSheet(SHEETS.TRAINERS);
  const trainerData = trainerSheet.getDataRange().getValues();
  for (let i = 1; i < trainerData.length; i++) {
    if (trainerData[i][0] === trainerID) {
      createNotification(trainerData[i][2], 'TRAINER', 'Batches Merged',
        `Batches ${sourceBatchIDs.join(', ')} have been merged into ${targetBatchName} (${targetBatchID})`, targetBatchID);
      break;
    }
  }
  
  logAudit('BATCH_MERGE', sourceBatchIDs.join(','), targetBatchID, targetBatchID, trainerID, classroomID, params.adminEmail, 'BATCH');
  logChangeHistory(params.adminEmail, 'BATCH_MERGE', {
    sourceBatches: sourceBatchIDs,
    targetBatch: targetBatchID,
    studentsMoved: movedStudents.length,
    newBatchName: targetBatchName
  });
  
  return jsonResponse(true, 'Batches merged successfully', {
    targetBatchID: targetBatchID,
    studentsMoved: movedStudents.length
  });
}

function handleAdminSplitBatch(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sourceBatchID = params.sourceBatchID;
  const splits = params.splits;
  
  const newBatchIDs = [];
  
  splits.forEach(split => {
    const newBatchID = generateID('BATCH');
    newBatchIDs.push(newBatchID);
    
    const batchSheet = getSheet(SHEETS.BATCHES);
    batchSheet.appendRow([
      newBatchID,
      split.name,
      split.trainerID,
      split.classroomID,
      split.mode || 'Offline',
      getCurrentDate(),
      params.endDate || '',
      '',
      'Active',
      'Split from: ' + sourceBatchID
    ]);
    
    const studentSheet = getSheet(SHEETS.STUDENTS);
    const studentData = studentSheet.getDataRange().getValues();
    
    for (let i = 1; i < studentData.length; i++) {
      if (split.studentIDs.includes(studentData[i][0])) {
        const row = i + 1;
        studentSheet.getRange(row, 5).setValue(newBatchID);
        
        createNotification(studentData[i][3], 'STUDENT', 'Batch Split',
          `You have been moved to a new batch: ${split.name} (${newBatchID})`, newBatchID);
      }
    }
    
    const moduleSheet = getSheet(SHEETS.MODULES);
    const moduleData = moduleSheet.getDataRange().getValues();
    
    for (let i = 1; i < moduleData.length; i++) {
      if (moduleData[i][1] === sourceBatchID && moduleData[i][3] !== 'COMPLETED') {
        const newModuleID = generateID('MOD');
        moduleSheet.appendRow([
          newModuleID,
          newBatchID,
          moduleData[i][2],
          'PENDING',
          moduleData[i][4],
          '',
          split.trainerID,
          'Split from ' + sourceBatchID
        ]);
      }
    }
  });
  
  const batchSheet = getSheet(SHEETS.BATCHES);
  const batchData = batchSheet.getDataRange().getValues();
  for (let i = 1; i < batchData.length; i++) {
    if (batchData[i][0] === sourceBatchID) {
      const row = i + 1;
      batchSheet.getRange(row, 9).setValue('Split');
      break;
    }
  }
  
  logAudit('BATCH_SPLIT', sourceBatchID, newBatchIDs.join(','), '', '', '', params.adminEmail, 'BATCH');
  logChangeHistory(params.adminEmail, 'BATCH_SPLIT', {
    sourceBatch: sourceBatchID,
    newBatches: newBatchIDs,
    splitCount: splits.length
  });
  
  return jsonResponse(true, 'Batch split successfully', { newBatchIDs: newBatchIDs });
}

function handleAdminChangeTrainer(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const batchID = params.batchID;
  const oldTrainerID = params.oldTrainerID;
  const newTrainerID = params.newTrainerID;
  
  const batchSheet = getSheet(SHEETS.BATCHES);
  const batchData = batchSheet.getDataRange().getValues();
  let batchName = '';
  
  for (let i = 1; i < batchData.length; i++) {
    if (batchData[i][0] === batchID) {
      const row = i + 1;
      batchSheet.getRange(row, 3).setValue(newTrainerID);
      batchName = batchData[i][1];
      break;
    }
  }
  
  const moduleSheet = getSheet(SHEETS.MODULES);
  const moduleData = moduleSheet.getDataRange().getValues();
  
  for (let i = 1; i < moduleData.length; i++) {
    if (moduleData[i][1] === batchID && moduleData[i][3] !== 'COMPLETED') {
      const row = i + 1;
      moduleSheet.getRange(row, 7).setValue(newTrainerID);
    }
  }
  
  const trainerSheet = getSheet(SHEETS.TRAINERS);
  const trainerData = trainerSheet.getDataRange().getValues();
  
  for (let i = 1; i < trainerData.length; i++) {
    if (trainerData[i][0] === oldTrainerID) {
      createNotification(trainerData[i][2], 'TRAINER', 'Batch Reassigned',
        `Batch ${batchName} (${batchID}) has been reassigned to another trainer`, batchID);
    }
    if (trainerData[i][0] === newTrainerID) {
      createNotification(trainerData[i][2], 'TRAINER', 'New Batch Assigned',
        `You have been assigned to batch ${batchName} (${batchID})`, batchID);
    }
  }
  
  const studentSheet = getSheet(SHEETS.STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][4] === batchID && studentData[i][6] === 'Active') {
      createNotification(studentData[i][3], 'STUDENT', 'Trainer Changed',
        `Your trainer has been changed for batch ${batchName}`, batchID);
    }
  }
  
  logAudit('TRAINER_CHANGE', oldTrainerID, newTrainerID, batchID, newTrainerID, '', params.adminEmail, 'BATCH');
  logChangeHistory(params.adminEmail, 'TRAINER_CHANGE', {
    batchID: batchID,
    oldTrainer: oldTrainerID,
    newTrainer: newTrainerID
  });
  
  return jsonResponse(true, 'Trainer changed successfully');
}

function handleAdminChangeClassroom(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const batchID = params.batchID;
  const oldClassroomID = params.oldClassroomID;
  const newClassroomID = params.newClassroomID;
  
  const batchSheet = getSheet(SHEETS.BATCHES);
  const batchData = batchSheet.getDataRange().getValues();
  let batchName = '';
  let trainerID = '';
  
  for (let i = 1; i < batchData.length; i++) {
    if (batchData[i][0] === batchID) {
      const row = i + 1;
      batchSheet.getRange(row, 4).setValue(newClassroomID);
      batchName = batchData[i][1];
      trainerID = batchData[i][2];
      break;
    }
  }
  
  let newClassroomName = newClassroomID;
  if (newClassroomID !== 'ONLINE') {
    const classroomSheet = getSheet(SHEETS.CLASSROOMS);
    const classroomData = classroomSheet.getDataRange().getValues();
    for (let i = 1; i < classroomData.length; i++) {
      if (classroomData[i][0] === newClassroomID) {
        newClassroomName = classroomData[i][1];
        break;
      }
    }
  }
  
  const studentSheet = getSheet(SHEETS.STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][4] === batchID && studentData[i][6] === 'Active') {
      createNotification(studentData[i][3], 'STUDENT', 'Classroom Changed',
        `Your classroom has been changed to ${newClassroomName} for batch ${batchName}`, batchID);
    }
  }
  
  const trainerSheet = getSheet(SHEETS.TRAINERS);
  const trainerData = trainerSheet.getDataRange().getValues();
  for (let i = 1; i < trainerData.length; i++) {
    if (trainerData[i][0] === trainerID) {
      createNotification(trainerData[i][2], 'TRAINER', 'Classroom Changed',
        `Classroom changed to ${newClassroomName} for batch ${batchName}`, batchID);
      break;
    }
  }
  
  logAudit('CLASSROOM_CHANGE', oldClassroomID, newClassroomID, batchID, trainerID, newClassroomID, params.adminEmail, 'BATCH');
  logChangeHistory(params.adminEmail, 'CLASSROOM_CHANGE', {
    batchID: batchID,
    oldClassroom: oldClassroomID,
    newClassroom: newClassroomID
  });
  
  return jsonResponse(true, 'Classroom changed successfully');
}

function handleAdminAddModule(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.MODULES);
  const moduleID = generateID('MOD');
  
  sheet.appendRow([
    moduleID,
    params.batchID,
    sanitizeInput(params.moduleName),
    'PENDING',
    params.tentativeDate || '',
    '',
    params.trainerID,
    params.notes || ''
  ]);
  
  logAudit('MODULE_ADDED', '', moduleID, params.batchID, params.trainerID, '', params.adminEmail, 'MODULE');
  
  return jsonResponse(true, 'Module added successfully', { moduleID: moduleID });
}

function handleAdminUpdateModule(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.MODULES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.moduleID) {
      const row = i + 1;
      
      if (params.moduleName) sheet.getRange(row, 3).setValue(sanitizeInput(params.moduleName));
      if (params.status) sheet.getRange(row, 4).setValue(params.status);
      if (params.tentativeDate) sheet.getRange(row, 5).setValue(params.tentativeDate);
      if (params.notes) sheet.getRange(row, 8).setValue(params.notes);
      
      logAudit('MODULE_UPDATED', '', params.moduleID, data[i][1], data[i][6], '', params.adminEmail, 'MODULE');
      
      return jsonResponse(true, 'Module updated successfully');
    }
  }
  
  return jsonResponse(false, 'Module not found');
}

function handleAdminAddTimingSlot(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.TIMING_SLOTS);
  const timingSlotID = generateID('TS');
  
  sheet.appendRow([
    timingSlotID,
    sanitizeInput(params.label),
    params.startTime,
    params.endTime,
    params.mode,
    params.adminEmail,
    getCurrentTimestamp()
  ]);
  
  logAudit('TIMING_SLOT_ADDED', '', timingSlotID, '', '', '', params.adminEmail, 'TIMING');
  
  return jsonResponse(true, 'Timing slot added successfully', { timingSlotID: timingSlotID });
}

function handleAdminGetAllBatches(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const batchSheet = getSheet(SHEETS.BATCHES);
  const batchData = batchSheet.getDataRange().getValues();
  
  const batches = [];
  for (let i = 1; i < batchData.length; i++) {
    const studentSheet = getSheet(SHEETS.STUDENTS);
    const studentData = studentSheet.getDataRange().getValues();
    let studentCount = 0;
    for (let j = 1; j < studentData.length; j++) {
      if (studentData[j][4] === batchData[i][0] && studentData[j][6] === 'Active') {
        studentCount++;
      }
    }
    
    const moduleSheet = getSheet(SHEETS.MODULES);
    const moduleData = moduleSheet.getDataRange().getValues();
    let totalModules = 0;
    let completedModules = 0;
    
    for (let j = 1; j < moduleData.length; j++) {
      if (moduleData[j][1] === batchData[i][0]) {
        totalModules++;
        if (moduleData[j][3] === 'COMPLETED') {
          completedModules++;
        }
      }
    }
    
    batches.push({
      batchID: batchData[i][0],
      name: batchData[i][1],
      trainerID: batchData[i][2],
      classroomID: batchData[i][3],
      mode: batchData[i][4],
      startDate: batchData[i][5],
      endDate: batchData[i][6],
      status: batchData[i][8],
      studentCount: studentCount,
      totalModules: totalModules,
      completedModules: completedModules
    });
  }
  
  return jsonResponse(true, 'Batches retrieved', { batches: batches });
}

function handleAdminGetAllStudents(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.STUDENTS);
  const data = sheet.getDataRange().getValues();
  
  const students = [];
  for (let i = 1; i < data.length; i++) {
    students.push({
      studentID: data[i][0],
      name: data[i][1],
      phone: data[i][2],
      email: data[i][3],
      batchID: data[i][4],
      status: data[i][6],
      registeredOn: data[i][7]
    });
  }
  
  return jsonResponse(true, 'Students retrieved', { students: students });
}

function handleAdminGetAllTrainers(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.TRAINERS);
  const data = sheet.getDataRange().getValues();
  
  const trainers = [];
  for (let i = 1; i < data.length; i++) {
    trainers.push({
      trainerID: data[i][0],
      name: data[i][1],
      email: data[i][2],
      phone: data[i][3],
      status: data[i][4]
    });
  }
  
  return jsonResponse(true, 'Trainers retrieved', { trainers: trainers });
}

function handleAdminGetAllClassrooms(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.CLASSROOMS);
  const data = sheet.getDataRange().getValues();
  
  const classrooms = [];
  for (let i = 1; i < data.length; i++) {
    classrooms.push({
      classroomID: data[i][0],
      locationName: data[i][1],
      latitude: data[i][2],
      longitude: data[i][3],
      radiusMeters: data[i][4],
      active: data[i][5] === 'Yes'
    });
  }
  
  return jsonResponse(true, 'Classrooms retrieved', { classrooms: classrooms });
}

function handleAdminGetAllModules(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.MODULES);
  const data = sheet.getDataRange().getValues();
  
  const modules = [];
  for (let i = 1; i < data.length; i++) {
    modules.push({
      moduleID: data[i][0],
      batchID: data[i][1],
      moduleName: data[i][2],
      status: data[i][3],
      tentativeDate: data[i][4],
      actualDate: data[i][5],
      trainerID: data[i][6],
      notes: data[i][7]
    });
  }
  
  return jsonResponse(true, 'Modules retrieved', { modules: modules });
}

function handleAdminAddClassroom(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.CLASSROOMS);
  const classroomID = generateID('CR');
  
  sheet.appendRow([
    classroomID,
    sanitizeInput(params.locationName),
    parseFloat(params.latitude),
    parseFloat(params.longitude),
    parseInt(params.radiusMeters),
    'Yes'
  ]);
  
  logAudit('CLASSROOM_ADDED', '', classroomID, '', '', classroomID, params.adminEmail, 'CLASSROOM');
  
  return jsonResponse(true, 'Classroom added successfully', { classroomID: classroomID });
}

function handleAdminGetAttendanceReport(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const batchID = params.batchID;
  const startDate = params.startDate;
  const endDate = params.endDate;
  
  const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
  const attendanceData = attendanceSheet.getDataRange().getValues();
  
  const report = [];
  for (let i = 1; i < attendanceData.length; i++) {
    const date = attendanceData[i][1];
    
    if ((batchID && attendanceData[i][3] !== batchID) ||
        (startDate && date < startDate) ||
        (endDate && date > endDate)) {
      continue;
    }
    
    report.push({
      date: date,
      studentID: attendanceData[i][2],
      batchID: attendanceData[i][3],
      punchIn: attendanceData[i][5],
      punchOut: attendanceData[i][6],
      duration: attendanceData[i][7],
      status: attendanceData[i][8]
    });
  }
  
  return jsonResponse(true, 'Attendance report generated', { report: report });
}

function handleAdminOverrideDevice(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.STUDENTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.studentID) {
      const row = i + 1;
      sheet.getRange(row, 6).setValue('');
      
      logAudit('DEVICE_OVERRIDE', data[i][5], 'RESET', '', '', '', params.adminEmail, 'STUDENT');
      
      return jsonResponse(true, 'Device binding reset. Student can login from new device.');
    }
  }
  
  return jsonResponse(false, 'Student not found');
}

function handleGetNotifications(params) {
  const email = params.email;
  const recipientType = params.recipientType;
  
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  const data = sheet.getDataRange().getValues();
  
  const notifications = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][2] === email && data[i][3] === recipientType) {
      notifications.push({
        notificationID: data[i][0],
        timestamp: data[i][1],
        subject: data[i][4],
        body: data[i][5],
        readStatus: data[i][7],
        batchID: data[i][8]
      });
      
      if (notifications.length >= 50) break;
    }
  }
  
  return jsonResponse(true, 'Notifications retrieved', { notifications: notifications });
}

function handleMarkNotificationRead(params) {
  const notificationID = params.notificationID;
  
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === notificationID) {
      const row = i + 1;
      sheet.getRange(row, 8).setValue('READ');
      return jsonResponse(true, 'Notification marked as read');
    }
  }
  
  return jsonResponse(false, 'Notification not found');
}

function handleSendNotifications(params) {
  if (!verifyAdmin(params.adminKey)) {
    return jsonResponse(false, 'Unauthorized');
  }
  
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  const data = sheet.getDataRange().getValues();
  
  let sentCount = 0;
  let failedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'PENDING') {
      const row = i + 1;
      const recipientEmail = data[i][2];
      const subject = data[i][4];
      const body = data[i][5];
      
      const emailBody = `
Digital Academy 360 Notification

${body}

---
This is an automated notification from DA360.
Please do not reply to this email.

Visit your portal: https://yoursite.github.io/da360
      `;
      
      const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    .header { background: #dc2626; color: white; padding: 20px; text-align: center; }
    .content { padding: 20px; }
    .footer { background: #f3f4f6; padding: 10px; text-align: center; font-size: 12px; color: #666; }
  </style>
</head>
<body>
  <div class="header">
    <h2>Digital Academy 360</h2>
  </div>
  <div class="content">
    <p>${body.replace(/\n/g, '<br>')}</p>
  </div>
  <div class="footer">
    <p>This is an automated notification from DA360.</p>
    <p>Please do not reply to this email.</p>
  </div>
</body>
</html>
      `;
      
      const sent = sendEmail(recipientEmail, subject, emailBody, htmlBody);
      
      if (sent) {
        sheet.getRange(row, 7).setValue('SENT');
        sentCount++;
      } else {
        sheet.getRange(row, 7).setValue('FAILED');
        failedCount++;
      }
    }
  }
  
  return jsonResponse(true, `Sent: ${sentCount}, Failed: ${failedCount}`);
}

// ==================== AUTOMATED TRIGGERS ====================

function autoMarkAbsent() {
  const today = getCurrentDate();
  const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
  const studentSheet = getSheet(SHEETS.STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  
  const activeStudents = [];
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][6] === 'Active') {
      activeStudents.push({
        studentID: studentData[i][0],
        batchID: studentData[i][4]
      });
    }
  }
  
  const attendanceData = attendanceSheet.getDataRange().getValues();
  const presentStudents = new Set();
  
  for (let i = 1; i < attendanceData.length; i++) {
    if (attendanceData[i][1] === today) {
      presentStudents.add(attendanceData[i][2]);
    }
  }
  
  activeStudents.forEach(student => {
    if (!presentStudents.has(student.studentID)) {
      const batchSheet = getSheet(SHEETS.BATCHES);
      const batchData = batchSheet.getDataRange().getValues();
      let trainerID = '';
      
      for (let i = 1; i < batchData.length; i++) {
        if (batchData[i][0] === student.batchID) {
          trainerID = batchData[i][2];
          break;
        }
      }
      
      attendanceSheet.appendRow([
        getCurrentTimestamp(),
        today,
        student.studentID,
        student.batchID,
        trainerID,
        '',
        '',
        0,
        'ABSENT',
        'No',
        'No',
        'Auto-marked absent'
      ]);
    }
  });
  
  Logger.log('Auto-marked absent for students without attendance');
}

function updateAttendanceMatrix() {
  const matrixSheet = getSheet(SHEETS.ATTENDANCE_MATRIX);
  const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
  const studentSheet = getSheet(SHEETS.STUDENTS);
  
  matrixSheet.clear();
  
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 30);
  
  const attendanceData = attendanceSheet.getDataRange().getValues();
  const dates = new Set();
  
  for (let i = 1; i < attendanceData.length; i++) {
    const date = attendanceData[i][1];
    if (date >= Utilities.formatDate(startDate, 'Asia/Kolkata', 'yyyy-MM-dd')) {
      dates.add(date);
    }
  }
  
  const dateArray = Array.from(dates).sort();
  
  const header = ['StudentID', 'Name', ...dateArray, 'Percentage'];
  matrixSheet.appendRow(header);
  
  const studentData = studentSheet.getDataRange().getValues();
  
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][6] === 'Active') {
      const studentID = studentData[i][0];
      const name = studentData[i][1];
      const row = [studentID, name];
      
      let totalDays = 0;
      let presentDays = 0;
      
      dateArray.forEach(date => {
        let status = '';
        for (let j = 1; j < attendanceData.length; j++) {
          if (attendanceData[j][2] === studentID && attendanceData[j][1] === date) {
            status = attendanceData[j][8];
            break;
          }
        }
        
        let mark = '';
        if (status === 'PRESENT') {
          mark = 'P';
          presentDays++;
          totalDays++;
        } else if (status === 'HALF_DAY') {
          mark = 'H';
          presentDays += 0.5;
          totalDays++;
        } else if (status === 'ABSENT') {
          mark = 'A';
          totalDays++;
        }
        
        row.push(mark);
      });
      
      const percentage = totalDays > 0 ? ((presentDays / totalDays) * 100).toFixed(2) + '%' : '0%';
      row.push(percentage);
      
      matrixSheet.appendRow(row);
    }
  }
  
  Logger.log('Attendance matrix updated');
}

function sendDailyTrainerSummary() {
  const today = getCurrentDate();
  const trainerSheet = getSheet(SHEETS.TRAINERS);
  const trainerData = trainerSheet.getDataRange().getValues();
  
  for (let i = 1; i < trainerData.length; i++) {
    if (trainerData[i][4] === 'Active') {
      const trainerID = trainerData[i][0];
      const trainerEmail = trainerData[i][2];
      const trainerName = trainerData[i][1];
      
      const batchSheet = getSheet(SHEETS.BATCHES);
      const batchData = batchSheet.getDataRange().getValues();
      
      let summary = `Daily Summary for ${trainerName}\nDate: ${today}\n\n`;
      
      for (let j = 1; j < batchData.length; j++) {
        if (batchData[j][2] === trainerID && batchData[j][8] === 'Active') {
          const batchID = batchData[j][0];
          const batchName = batchData[j][1];
          
          const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
          const attendanceData = attendanceSheet.getDataRange().getValues();
          
          let present = 0;
          let halfDay = 0;
          let absent = 0;
          
          for (let k = 1; k < attendanceData.length; k++) {
            if (attendanceData[k][1] === today && attendanceData[k][3] === batchID) {
              const status = attendanceData[k][8];
              if (status === 'PRESENT') present++;
              else if (status === 'HALF_DAY') halfDay++;
              else if (status === 'ABSENT') absent++;
            }
          }
          
          summary += `Batch: ${batchName}\n`;
          summary += `  Present: ${present}\n`;
          summary += `  Half Day: ${halfDay}\n`;
          summary += `  Absent: ${absent}\n\n`;
        }
      }
      
      if (summary.includes('Batch:')) {
        sendEmail(trainerEmail, 'Daily Attendance Summary', summary);
      }
    }
  }
  
  Logger.log('Daily trainer summaries sent');
}

// ==================== SETUP FUNCTIONS ====================

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  ScriptApp.newTrigger('autoMarkAbsent')
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .create();
  
  ScriptApp.newTrigger('updateAttendanceMatrix')
    .timeBased()
    .everyHours(1)
    .create();
  
  ScriptApp.newTrigger('sendDailyTrainerSummary')
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .create();
  
  Logger.log('Triggers set up successfully');
}

function initializeSheets() {
  const ss = getSpreadsheet();
  
  Object.values(SHEETS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log(`Created sheet: ${sheetName}`);
    }
  });
  
  Logger.log('All sheets initialized');
}

// ==================== TEST FUNCTION ====================

function testAPI() {
  const testStudent = {
    action: 'loginStudent',
    studentID: 'STD001',
    deviceInfo: 'test-device-info'
  };
  
  const result = doPost({
    postData: {
      contents: JSON.stringify(testStudent)
    }
  });
  
  Logger.log(result.getContent());
}

function testAttendanceCode() {
  const studentID = 'STD001';
  const date = getCurrentDate();
  const code = generateDailyAttendanceCode(studentID, date);
  Logger.log(`Generated code for ${studentID} on ${date}: ${code}`);
  
  const isValid = validateAttendanceCode(studentID, code);
  Logger.log(`Code validation: ${isValid}`);
}
/**
 * Digital Academy 360 - Google Apps Script Backend
 * Complete API and Automation Engine - FINAL VERSION
 * Version: 1.0 - Production Ready
 */

// ==================== CONFIGURATION ====================

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // Replace with actual ID
const ADMIN_SECRET_KEY = 'DA360_ADMIN_SECRET_2025'; // Change this!

// Sheet Names
const SHEETS = {
  STUDENTS: 'STUDENTS',
  TRAINERS: 'TRAINERS',
  BATCHES: 'BATCHES',
  CLASSROOMS: 'CLASSROOMS',
  MODULES: 'MODULES',
  ATTENDANCE_LOG: 'ATTENDANCE_LOG',
  TRAINER_ATTENDANCE: 'TRAINER_ATTENDANCE',
  AUDIT_LOG: 'AUDIT_LOG',
  CHANGE_HISTORY: 'CHANGE_HISTORY',
  NOTIFICATIONS: 'NOTIFICATIONS',
  CONFIG: 'CONFIG',
  ATTENDANCE_MATRIX: 'ATTENDANCE_MATRIX',
  TIMING_SLOTS: 'TIMING_SLOTS',
  BATCH_TIMINGS: 'BATCH_TIMINGS'
};

// ==================== UTILITY FUNCTIONS ====================

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

function getCurrentTimestamp() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
}

function getCurrentDate() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
}

function getCurrentTime() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'HH:mm:ss');
}

function generateID(prefix) {
  return prefix + Date.now() + Math.random().toString(36).substr(2, 5).toUpperCase();
}

function hashDevice(deviceInfo) {
  const signature = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, deviceInfo)
  );
  return signature;
}

function sanitizeInput(input) {
  if (typeof input === 'string') {
    return input.trim().replace(/[<>]/g, '');
  }
  return input;
}

function sendEmail(to, subject, body, htmlBody) {
  try {
    if (htmlBody) {
      MailApp.sendEmail({
        to: to,
        subject: subject,
        body: body,
        htmlBody: htmlBody
      });
    } else {
      MailApp.sendEmail(to, subject, body);
    }
    return true;
  } catch (e) {
    Logger.log('Email Error: ' + e.message);
    return false;
  }
}

function getConfigValue(key) {
  const sheet = getSheet(SHEETS.CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

function logAudit(action, oldValue, newValue, batchID, trainerID, classroomID, adminID, entity) {
  const sheet = getSheet(SHEETS.AUDIT_LOG);
  sheet.appendRow([
    getCurrentTimestamp(),
    action,
    oldValue || '',
    newValue || '',
    batchID || '',
    trainerID || '',
    classroomID || '',
    adminID || '',
    entity || ''
  ]);
}

function logChangeHistory(actorEmail, actionType, detailsJSON) {
  const sheet = getSheet(SHEETS.CHANGE_HISTORY);
  sheet.appendRow([
    getCurrentTimestamp(),
    actorEmail,
    actionType,
    JSON.stringify(detailsJSON)
  ]);
}

function createNotification(recipientEmail, recipientType, subject, body, batchID) {
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  const notifID = generateID('NOTIF');
  sheet.appendRow([
    notifID,
    getCurrentTimestamp(),
    recipientEmail,
    recipientType,
    subject,
    body,
    'PENDING',
    'UNREAD',
    batchID || ''
  ]);
  return notifID;
}

// ==================== GEOLOCATION ====================

function haversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000; // Earth radius in meters
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function validateGeofence(studentLat, studentLon, classroomID) {
  const sheet = getSheet(SHEETS.CLASSROOMS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === classroomID) {
      const classLat = data[i][2];
      const classLon = data[i][3];
      const radius = data[i][4];
      
      const distance = haversineDistance(studentLat, studentLon, classLat, classLon);
      return distance <= radius;
    }
  }
  return false;
}

// ==================== ATTENDANCE CODE GENERATION ====================

function generateDailyAttendanceCode(studentID, date) {
  const seed = getConfigValue('DailyCodeSeed') || 'DA360SECRET2025';
  const rawString = studentID + date + seed;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawString);
  const hashHex = hash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
  const code = parseInt(hashHex.substr(0, 8), 16) % 10000;
  return code.toString().padStart(4, '0');
}

function validateAttendanceCode(studentID, providedCode) {
  const today = getCurrentDate();
  const validCode = generateDailyAttendanceCode(studentID, today);
  return providedCode === validCode;
}

// ==================== WEB APP ENTRY POINT ====================

function doGet(e) {
  // Add no-cache headers
  const output = ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'DA360 API is running',
    version: '1.0',
    timestamp: getCurrentTimestamp()
  }));
  
  output.setMimeType(ContentService.MimeType.JSON);
  
  return output;
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    
    let response;
    
    // Route to appropriate handler
    switch(action) {
      // Student endpoints
      case 'loginStudent':
        response = handleLoginStudent(params);
        break;
      case 'getStudentDashboard':
        response = handleGetStudentDashboard(params);
        break;
      case 'getAttendanceCode':
        response = handleGetAttendanceCode(params);
        break;
      case 'punchIn':
        response = handlePunchIn(params);
        break;
      case 'punchOut':
        response = handlePunchOut(params);
        break;
      case 'getStudentAttendance':
        response = handleGetStudentAttendance(params);
        break;
      case 'getNotifications':
        response = handleGetNotifications(params);
        break;
      case 'markNotificationRead':
        response = handleMarkNotificationRead(params);
        break;
      
      // Trainer endpoints
      case 'loginTrainer':
        response = handleLoginTrainer(params);
        break;
      case 'getTrainerDashboard':
        response = handleGetTrainerDashboard(params);
        break;
      case 'getTrainerBatches':
        response = handleGetTrainerBatches(params);
        break;
      case 'markTrainerAbsent':
        response = handleMarkTrainerAbsent(params);
        break;
      case 'markModuleCompleted':
        response = handleMarkModuleCompleted(params);
        break;
      case 'rescheduleModule':
        response = handleRescheduleModule(params);
        break;
      
      // Admin endpoints
      case 'adminAddStudent':
        response = handleAdminAddStudent(params);
        break;
      case 'adminUpdateStudent':
        response = handleAdminUpdateStudent(params);
        break;
      case 'adminMoveStudent':
        response = handleAdminMoveStudent(params);
        break;
      case 'adminAddTrainer':
        response = handleAdminAddTrainer(params);
        break;
      case 'adminAddBatch':
        response = handleAdminAddBatch(params);
        break;
      case 'adminUpdateBatch':
        response = handleAdminUpdateBatch(params);
        break;
      case 'adminMergeBatch':
        response = handleAdminMergeBatch(params);
        break;
      case 'adminSplitBatch':
        response = handleAdminSplitBatch(params);
        break;
      case 'adminChangeTrainer':
        response = handleAdminChangeTrainer(params);
        break;
      case 'adminChangeClassroom':
        response = handleAdminChangeClassroom(params);
        break;
      case 'adminAddModule':
        response = handleAdminAddModule(params);
        break;
      case 'adminUpdateModule':
        response = handleAdminUpdateModule(params);
        break;
      case 'adminAddTimingSlot':
        response = handleAdminAddTimingSlot(params);
        break;
      case 'adminGetAllBatches':
        response = handleAdminGetAllBatches(params);
        break;
      case 'adminGetAllStudents':
        response = handleAdminGetAllStudents(params);
        break;
      case 'adminGetAllTrainers':
        response = handleAdminGetAllTrainers(params);
        break;
      case 'adminGetAllClassrooms':
        response = handleAdminGetAllClassrooms(params);
        break;
      case 'adminGetAllModules':
        response = handleAdminGetAllModules(params);
        break;
      case 'adminAddClassroom':
        response = handleAdminAddClassroom(params);
        break;
      case 'adminGetAttendanceReport':
        response = handleAdminGetAttendanceReport(params);
        break;
      case 'adminOverrideDevice':
        response = handleAdminOverrideDevice(params);
        break;
      case 'sendNotifications':
        response = handleSendNotifications(params);
        break;
      
      default:
        response = jsonResponse(false, 'Invalid action: ' + action);
    }
    
    return response;
  } catch (e) {
    Logger.log('Error: ' + e.message + ' | Stack: ' + e.stack);
    return jsonResponse(false, 'Server error: ' + e.message);
  }
}

function jsonResponse(success, message, data) {
  const response = { 
    success: success, 
    message: message, 
    data: data || null,
    timestamp: getCurrentTimestamp()
  };
  
  const output = ContentService.createTextOutput(JSON.stringify(response));
  output.setMimeType(ContentService.MimeType.JSON);
  
  return output;
}

// ==================== STUDENT ENDPOINTS ====================

function handleLoginStudent(params) {
  const studentID = sanitizeInput(params.studentID);
  const deviceInfo = params.deviceInfo;
  
  const sheet = getSheet(SHEETS.STUDENTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === studentID) {
      const storedDevice = data[i][5];
      const status = data[i][6];
      
      if (status !== 'Active') {
        return jsonResponse(false, 'Account is inactive');
      }
      
      const deviceHash = hashDevice(deviceInfo);
      
      // First time login
      if (!storedDevice) {
        sheet.getRange(i + 1, 6).setValue(deviceHash);
        return jsonResponse(true, 'Login successful - Device registered', {
          studentID: studentID,
          name: data[i][1],
          email: data[i][3],
          batchID: data[i][4],
          isFirstLogin: true
        });
      }
      
      // Verify device
      if (storedDevice !== deviceHash) {
        return jsonResponse(false, 'Device mismatch. Please contact admin.');
      }
      
      return jsonResponse(true, 'Login successful', {
        studentID: studentID,
        name: data[i][1],
        email: data[i][3],
        batchID: data[i][4]
      });
    }
  }
  
  return jsonResponse(false, 'Invalid Student ID');
}

function handleGetStudentDashboard(params) {
  const studentID = params.studentID;
  
  const studentSheet = getSheet(SHEETS.STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  
  let student = null;
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][0] === studentID) {
      student = {
        studentID: studentData[i][0],
        name: studentData[i][1],
        email: studentData[i][3],
        batchID: studentData[i][4]
      };
      break;
    }
  }
  
  if (!student) {
    return jsonResponse(false, 'Student not found');
  }
  
  // Get batch details
  const batchSheet = getSheet(SHEETS.BATCHES);
  const batchData = batchSheet.getDataRange().getValues();
  let batch = null;
  
  for (let i = 1; i < batchData.length; i++) {
    if (batchData[i][0] === student.batchID) {
      batch = {
        batchID: batchData[i][0],
        name: batchData[i][1],
        trainerID: batchData[i][2],
        classroomID: batchData[i][3],
        mode: batchData[i][4]
      };
      break;
    }
  }
  
  // Get trainer name
  let trainerName = '';
  if (batch) {
    const trainerSheet = getSheet(SHEETS.TRAINERS);
    const trainerData = trainerSheet.getDataRange().getValues();
    for (let i = 1; i < trainerData.length; i++) {
      if (trainerData[i][0] === batch.trainerID) {
        trainerName = trainerData[i][1];
        break;
      }
    }
  }
  
  // Get classroom name
  let classroomName = batch ? batch.classroomID : '';
  if (batch && batch.classroomID !== 'ONLINE') {
    const classroomSheet = getSheet(SHEETS.CLASSROOMS);
    const classroomData = classroomSheet.getDataRange().getValues();
    for (let i = 1; i < classroomData.length; i++) {
      if (classroomData[i][0] === batch.classroomID) {
        classroomName = classroomData[i][1];
        break;
      }
    }
  }
  
  // Calculate attendance percentage
  const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
  const attendanceData = attendanceSheet.getDataRange().getValues();
  let totalDays = 0;
  let presentDays = 0;
  
  for (let i = 1; i < attendanceData.length; i++) {
    if (attendanceData[i][2] === studentID) {
      totalDays++;
      if (attendanceData[i][8] === 'PRESENT') {
        presentDays++;
      } else if (attendanceData[i][8] === 'HALF_DAY') {
        presentDays += 0.5;
      }
    }
  }
  
  const attendancePercentage = totalDays > 0 ? ((presentDays / totalDays) * 100).toFixed(2) : 0;
  
  // Get current module
  const moduleSheet = getSheet(SHEETS.MODULES);
  const moduleData = moduleSheet.getDataRange().getValues();
  let currentModule = '';
  
  for (let i = 1; i < moduleData.length; i++) {
    if (moduleData[i][1] === student.batchID && moduleData[i][3] === 'IN_PROGRESS') {
      currentModule = moduleData[i][2];
      break;
    }
  }
  
  return jsonResponse(true, 'Dashboard loaded', {
    student: student,
    batch: batch,
    trainerName: trainerName,
    classroomName: classroomName,
    attendancePercentage: attendancePercentage,
    currentModule: currentModule
  });
}

function handleGetAttendanceCode(params) {
  const studentID = params.studentID;
  const today = getCurrentDate();
  const code = generateDailyAttendanceCode(studentID, today);
  
  return jsonResponse(true, 'Code generated', { code: code });
}

function handlePunchIn(params) {
  const studentID = sanitizeInput(params.studentID);
  const attendanceCode = sanitizeInput(params.code);
  const latitude = parseFloat(params.latitude);
  const longitude = parseFloat(params.longitude);
  const deviceInfo = params.deviceInfo;
  
  // Validate attendance code
  if (!validateAttendanceCode(studentID, attendanceCode)) {
    return jsonResponse(false, 'Invalid attendance code');
  }
  
  // Verify device
  const studentSheet = getSheet(SHEETS.STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  let student = null;
  
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][0] === studentID) {
      const storedDevice = studentData[i][5];
      const deviceHash = hashDevice(deviceInfo);
      
      if (storedDevice !== deviceHash) {
        return jsonResponse(false, 'Device verification failed');
      }
      
      student = {
        studentID: studentData[i][0],
        batchID: studentData[i][4]
      };
      break;
    }
  }
  
  if (!student) {
    return jsonResponse(false, 'Student not found');
  }
  
  // Get batch details
  const batchSheet = getSheet(SHEETS.BATCHES);
  const batchData = batchSheet.getDataRange().getValues();
  let batch = null;
  
  for (let i = 1; i < batchData.length; i++) {
    if (batchData[i][0] === student.batchID) {
      batch = {
        batchID: batchData[i][0],
        trainerID: batchData[i][2],
        classroomID: batchData[i][3],
        mode: batchData[i][4]
      };
      break;
    }
  }
  
  // Validate geofence for offline batches
  let geoVerified = 'N/A';
  if (batch.mode === 'Offline') {
    const isWithinRange = validateGeofence(latitude, longitude, batch.classroomID);
    if (!isWithinRange) {
      return jsonResponse(false, 'You are not within the classroom geofence');
    }
    geoVerified = 'Yes';
  }
  
  // Check if already punched in today
  const today = getCurrentDate();
  const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
  const attendanceData = attendanceSheet.getDataRange().getValues();
  
  for (let i = 1; i < attendanceData.length; i++) {
    if (attendanceData[i][1] === today && 
        attendanceData[i][2] === studentID && 
        attendanceData[i][5]) {
      return jsonResponse(false, 'Already punched in today');
    }
  }
  
  // Record punch in
  const timestamp = getCurrentTimestamp();
  const time = getCurrentTime();
  
  attendanceSheet.appendRow([
    timestamp,
    today,
    studentID,
    student.batchID,
    batch.trainerID,
    time,
    '',
    0,
    'PENDING',
    'Yes',
    geoVerified,
    'Punch in recorded'
  ]);
  
  return jsonResponse(true, 'Punch in successful', { punchInTime: time });
}

function handlePunchOut(params) {
  const studentID = sanitizeInput(params.studentID);
  const today = getCurrentDate();
  
  const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
  const attendanceData = attendanceSheet.getDataRange().getValues();
  
  for (let i = attendanceData.length - 1; i >= 1; i--) {
    if (attendanceData[i][1] === today && attendanceData[i][2] === studentID) {
      const punchInTime = attendanceData[i][5];
      
      if (!punchInTime) {
        return jsonResponse(false, 'No punch in record found for today');
      }
      
      const punchOutTime = attendanceData[i][6];
      if (punchOutTime) {
        return jsonResponse(false, 'Already punched out today');
      }
      
      const currentTime = getCurrentTime();
      const duration = calculateTimeDifferenceInMinutes(punchInTime, currentTime);
      
      if (duration < 0) {
        return jsonResponse(false, 'Invalid time calculation. Please contact admin.');
      }
      
      const minDuration = parseInt(getConfigValue('MinDurationMinutes')) || 90;
      const halfDayThreshold = parseInt(getConfigValue('HalfDayThreshold')) || 60;
      
      let status = 'ABSENT';
      let message = '';
      
      if (duration >= minDuration) {
        status = 'PRESENT';
        message = `Attendance marked PRESENT (${duration} minutes)`;
      } else if (duration >= halfDayThreshold) {
        status = 'HALF_DAY';
        message = `Marked HALF DAY (${duration} minutes). Need ${minDuration} min for full attendance.`;
      } else {
        status = 'ABSENT';
        message = `Marked ABSENT (${duration} minutes). Need ${halfDayThreshold} min for half day.`;
      }
      
      const row = i + 1;
      attendanceSheet.getRange(row, 7).setValue(currentTime);
      attendanceSheet.getRange(row, 8).setValue(duration);
      attendanceSheet.getRange(row, 9).setValue(status);
      attendanceSheet.getRange(row, 12).setValue(message);
      
      return jsonResponse(true, 'Punch out successful', {
        punchOutTime: currentTime,
        duration: duration,
        status: status,
        message: message
      });
    }
  }
  
  return jsonResponse(false, 'No punch in record found for today');
}

function calculateTimeDifferenceInMinutes(startTime, endTime) {
  try {
    const startParts = startTime.split(':');
    const endParts = endTime.split(':');
    
    const startMinutes = parseInt(startParts[0]) * 60 + parseInt(startParts[1]);
    const endMinutes = parseInt(endParts[0]) * 60 + parseInt(endParts[1]);
    
    let duration = endMinutes - startMinutes;
    
    if (duration < 0) {
      duration = (24 * 60) + duration;
    }
    
    return duration;
  } catch (e) {
    Logger.log('Time calculation error: ' + e.message);
    return 0;
  }
}

function handleGetStudentAttendance(params) {
  const studentID = params.studentID;
  const month = params.month || getCurrentDate().substr(0, 7);
  
  const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
  const attendanceData = attendanceSheet.getDataRange().getValues();
  
  const records = [];
  for (let i = 1; i < attendanceData.length; i++) {
    if (attendanceData[i][2] === studentID && attendanceData[i][1].toString().startsWith(month)) {
      records.push({
        date: attendanceData[i][1],
        punchIn: attendanceData[i][5],
        punchOut: attendanceData[i][6],
        duration: attendanceData[i][7],
        status: attendanceData[i][8]
      });
    }
  }
  
  return jsonResponse(true, 'Attendance records retrieved', { records: records });
}

// ==================== TRAINER ENDPOINTS ====================

function handleLoginTrainer(params) {
  const trainerID = sanitizeInput(params.trainerID);
  
  const sheet = getSheet(SHEETS.TRAINERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === trainerID) {
      const status = data[i][4];
      
      if (status !== 'Active') {
        return jsonResponse(false, 'Account is not active');
      }
      
      return jsonResponse(true, 'Login successful', {
        trainerID: trainerID,
        name: data[i][1],
        email: data[i][2]
      });
    }
  }
  
  return jsonResponse(false, 'Invalid Trainer ID');
}

function handleGetTrainerDashboard(params) {
  const trainerID = params.trainerID;
  
  const batchSheet = getSheet(SHEETS.BATCHES);
  const batchData = batchSheet.getDataRange().getValues();
  
  const batches = [];
  for (let i = 1; i < batchData.length; i++) {
    if (batchData[i][2] === trainerID && batchData[i][8] === 'Active') {
      const batchID = batchData[i][0];
      
      let timing = '';
      const timingSheet = getSheet(SHEETS.BATCH_TIMINGS);
      const timingData = timingSheet.getDataRange().getValues();
      for (let j = 1; j < timingData.length; j++) {
        if (timingData[j][1] === batchID && timingData[j][5] === 'Active') {
          const slotID = timingData[j][2];
          const slotSheet = getSheet(SHEETS.TIMING_SLOTS);
          const slotData = slotSheet.getDataRange().getValues();
          for (let k = 1; k < slotData.length; k++) {
            if (slotData[k][0] === slotID) {
              timing = slotData[k][2] + ' - ' + slotData[k][3];
              break;
            }
          }
          break;
        }
      }
      
      const studentSheet = getSheet(SHEETS.STUDENTS);
      const studentData = studentSheet.getDataRange().getValues();
      let studentCount = 0;
      for (let j = 1; j < studentData.length; j++) {
        if (studentData[j][4] === batchID && studentData[j][6] === 'Active') {
          studentCount++;
        }
      }
      
      batches.push({
        batchID: batchID,
        name: batchData[i][1],
        classroom: batchData[i][3],
        mode: batchData[i][4],
        timing: timing || 'Not set',
        studentCount: studentCount
      });
    }
  }
  
  return jsonResponse(true, 'Dashboard loaded', { batches: batches });
}

function handleGetTrainerBatches(params) {
  const trainerID = params.trainerID;
  const batchID = params.batchID;
  
  const studentSheet = getSheet(SHEETS.STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  
  const students = [];
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][4] === batchID && studentData[i][6] === 'Active') {
      const today = getCurrentDate();
      const attendanceSheet = getSheet(SHEETS.ATTENDANCE_LOG);
      const attendanceData = attendanceSheet.getDataRange().getValues();
      
      let todayStatus = 'NOT_MARKED';
      for (let j = 1; j < attendanceData.length; j++) {
        if (attendanceData[j][1] === today && attendanceData[j][2] === studentData[i][0]) {
          todayStatus = attendanceData[j][8];
          break;
        }
      }
      
      students.push({
        studentID: studentData[i][0],
        name: studentData[i][1],
        email: studentData[i][3],
        todayStatus: todayStatus
      });
    }
  }
  
  const moduleSheet = getSheet(SHEETS.MODULES);
  const moduleData = moduleSheet.getDataRange().getValues();
  
  const modules = [];
  for (let i = 1; i < moduleData.length; i++) {
    if (moduleData[i][1] === batchID) {
      modules.push({
        moduleID: moduleData[i][0],
        moduleName: moduleData[i][2],
        status: moduleData[i][3],
        tentativeDate: moduleData[i][4],
        actualDate: moduleData[i][5],
        notes: moduleData[i][7]
      });
    }
  }
  
  return jsonResponse(true, 'Batch details loaded', {
    students: students,
    modules: modules
  });
}

function handleMarkTrainerAbsent(params) {
  const trainerID = params.trainerID;
  const date = params.date || getCurrentDate();
  const remarks = params.remarks || 'Trainer marked absent';
  
  const sheet = getSheet(SHEETS.TRAINER_ATTENDANCE);
  sheet.appendRow([date, trainerID, 'ABSENT', '', remarks]);
  
  const adminEmail = getConfigValue('AdminEmail');
  createNotification(adminEmail, 'ADMIN', 'Trainer Absent', 
    `Trainer ${trainerID} marked absent on ${date}. Please assign substitute.`, '');
  
  logAudit('TRAINER_ABSENT', '', trainerID, '', trainerID, '', 'system', 'TRAINER');
  
  return jsonResponse(true, 'Marked as absent. Admin notified.');
}

function handleMarkModuleCompleted(params) {
  const moduleID = params.moduleID;
  const trainerID = params.trainerID;
  const notes = params.notes || '';
  
  const sheet = getSheet(SHEETS.MODULES);
  const data = sheet.getData
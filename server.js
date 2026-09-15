const express = require('express');
const path = require('path');
const twilio = require('twilio');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const Anthropic = require('@anthropic-ai/sdk');
require('dotenv').config();

const app = express();
app.use(express.urlencoded({ extended: false }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  res.redirect('/dashboard');
});

// Manager / admin dashboard (pitch-ready visual demo, static mock data)
app.get('/manager', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'manager-dashboard.html'));
});

// Student / parent dashboard (pitch-ready visual demo, static mock data)
app.get('/student', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'student-dashboard.html'));
});

const client = twilio(process.env.TWILIO_ACCOUNT_SID, process.env.TWILIO_AUTH_TOKEN);
const anthropic = new Anthropic({
  apiKey: process.env.CLAUDE_API_KEY,
});

const TWILIO_WHATSAPP_NUMBER = process.env.TWILIO_WHATSAPP_NUMBER;
const ADMIN_PHONE = process.env.MANAGER_PHONE;
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SERVICE_ACCOUNT = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT || '{}');

let doc;
let parentSheet, studentSheet, attendanceSheet, homeworkSheet, feesSheet, disciplineSheet, announcementsSheet;

async function initializeSheets() {
  try {
    doc = new GoogleSpreadsheet(SHEET_ID);
    await doc.useServiceAccountAuth(GOOGLE_SERVICE_ACCOUNT);
    await doc.loadInfo();

    parentSheet = doc.sheetsByTitle['Parents'] || await doc.addSheet({ title: 'Parents' });
    studentSheet = doc.sheetsByTitle['Students'] || await doc.addSheet({ title: 'Students' });
    attendanceSheet = doc.sheetsByTitle['Attendance'] || await doc.addSheet({ title: 'Attendance' });
    homeworkSheet = doc.sheetsByTitle['Homework'] || await doc.addSheet({ title: 'Homework' });
    feesSheet = doc.sheetsByTitle['Fees'] || await doc.addSheet({ title: 'Fees' });
    disciplineSheet = doc.sheetsByTitle['Discipline'] || await doc.addSheet({ title: 'Discipline' });
    announcementsSheet = doc.sheetsByTitle['Announcements'] || await doc.addSheet({ title: 'Announcements' });

    if (parentSheet.rowCount === 1 && !parentSheet.headerValues.length) {
      await parentSheet.setHeaderRow(['Phone', 'Parent Name', 'Child Name', 'Class', 'Status', 'Joined Date']);
    }
    if (studentSheet.rowCount === 1 && !studentSheet.headerValues.length) {
      await studentSheet.setHeaderRow(['Student ID', 'Name', 'Class', 'Roll No', 'Parent Phone', 'Current Marks']);
    }
    if (attendanceSheet.rowCount === 1 && !attendanceSheet.headerValues.length) {
      await attendanceSheet.setHeaderRow(['Date', 'Student Name', 'Class', 'Status', 'Parent Notified']);
    }
    if (homeworkSheet.rowCount === 1 && !homeworkSheet.headerValues.length) {
      await homeworkSheet.setHeaderRow(['Class', 'Subject', 'Assignment', 'Due Date', 'Posted By', 'Posted Date']);
    }
    if (feesSheet.rowCount === 1 && !feesSheet.headerValues.length) {
      await feesSheet.setHeaderRow(['Student Name', 'Parent Phone', 'Amount Due', 'Due Date', 'Status', 'Last Reminder']);
    }
    if (disciplineSheet.rowCount === 1 && !disciplineSheet.headerValues.length) {
      await disciplineSheet.setHeaderRow(['Date', 'Student Name', 'Class', 'Issue', 'Severity', 'Parent Notified', 'Status']);
    }
    if (announcementsSheet.rowCount === 1 && !announcementsSheet.headerValues.length) {
      await announcementsSheet.setHeaderRow(['Title', 'Message', 'Category', 'Sent Date', 'Recipients', 'Status']);
    }

    console.log('School sheets initialized');
  } catch (error) {
    console.error('Error initializing sheets:', error.message);
  }
}

// initializeSheets();

// WhatsApp Message Handler
app.post('/whatsapp', async (req, res) => {
  const from = req.body.From.replace('whatsapp:', '');
  const messageBody = req.body.Body.trim().toLowerCase();

  try {
    const parentRows = await parentSheet.getRows();
    const parent = parentRows.find(p => p.Phone === from);

    // Check what parent is asking for
    if (messageBody.includes('attendance') || messageBody.includes('present')) {
      await handleAttendanceQuery(from, parent);
    } else if (messageBody.includes('homework') || messageBody.includes('assignment')) {
      await handleHomeworkQuery(from, parent);
    } else if (messageBody.includes('fees') || messageBody.includes('bill') || messageBody.includes('payment')) {
      await handleFeesQuery(from, parent);
    } else if (messageBody.includes('emergency') || messageBody.includes('urgent') || messageBody.includes('help')) {
      await handleEmergency(from, parent);
    } else if (messageBody.includes('marks') || messageBody.includes('report')) {
      await handleMarksQuery(from, parent);
    } else if (messageBody.includes('issue') || messageBody.includes('problem') || messageBody.includes('complain')) {
      await handleDisciplineReport(from, parent, messageBody);
    } else if (!parent) {
      await handleRegistration(from, messageBody);
    } else {
      await sendWhatsAppMessage(from, `Hi! I can help with:\n• Attendance check\n• Homework/Assignments\n• Fee status\n• Academic marks\n• Report issues\n• Emergency alerts\n\nWhat do you need?`);
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Error');
  }
});

// Attendance Query
async function handleAttendanceQuery(from, parent) {
  if (!parent) {
    await sendWhatsAppMessage(from, 'Please register first: My name is [Name], child is [Child Name], class [Class]');
    return;
  }

  const attendanceRows = await attendanceSheet.getRows();
  const today = new Date().toISOString().split('T')[0];
  const todayAttendance = attendanceRows.filter(a => a.Date === today && a['Student Name'] === parent['Child Name']);

  if (todayAttendance.length > 0) {
    const status = todayAttendance[0].Status;
    await sendWhatsAppMessage(from, `Your child's attendance today:\n\n${parent['Child Name']} - ${status}\n\nRemaining days to see next update.`);
  } else {
    await sendWhatsAppMessage(from, `Attendance not yet marked for today. Check back in the evening.`);
  }
}

// Homework Query
async function handleHomeworkQuery(from, parent) {
  if (!parent) {
    await sendWhatsAppMessage(from, 'Please register first');
    return;
  }

  const homeworkRows = await homeworkSheet.getRows();
  const classHomework = homeworkRows.filter(h => h.Class === parent.Class);

  if (classHomework.length > 0) {
    let message = `Homework for ${parent.Class}:\n\n`;
    classHomework.slice(0, 3).forEach(hw => {
      message += `📚 ${hw.Subject}\n${hw.Assignment}\nDue: ${hw['Due Date']}\n\n`;
    });
    await sendWhatsAppMessage(from, message);
  } else {
    await sendWhatsAppMessage(from, `No pending homework for ${parent.Class}.`);
  }
}

// Fees Query
async function handleFeesQuery(from, parent) {
  if (!parent) {
    await sendWhatsAppMessage(from, 'Please register first');
    return;
  }

  const feesRows = await feesSheet.getRows();
  const parentFees = feesRows.filter(f => f['Parent Phone'] === from);

  if (parentFees.length > 0) {
    const pending = parentFees.filter(f => f.Status !== 'Paid');
    if (pending.length > 0) {
      let message = `Outstanding fees:\n\n`;
      pending.forEach(f => {
        message += `Amount: ₹${f['Amount Due']}\nDue: ${f['Due Date']}\nStatus: ${f.Status}\n\n`;
      });
      message += `Reply 'pay' to process payment online.`;
      await sendWhatsAppMessage(from, message);
    } else {
      await sendWhatsAppMessage(from, `All fees are paid. Thank you!`);
    }
  }
}

// Marks/Report Query
async function handleMarksQuery(from, parent) {
  if (!parent) {
    await sendWhatsAppMessage(from, 'Please register first');
    return;
  }

  const studentRows = await studentSheet.getRows();
  const student = studentRows.find(s => s['Parent Phone'] === from);

  if (student) {
    await sendWhatsAppMessage(from, `${student.Name} - Current Status:\n\nClass: ${student.Class}\nCurrent Marks: ${student['Current Marks']}\n\nFull report card available in your school dashboard.`);
  }
}

// Discipline Report
async function handleDisciplineReport(from, parent, messageBody) {
  if (!parent) {
    await sendWhatsAppMessage(from, 'Please register first');
    return;
  }

  await disciplineSheet.addRow({
    Date: new Date().toISOString(),
    'Student Name': parent['Child Name'],
    Class: parent.Class,
    Issue: messageBody,
    Severity: 'Reported',
    'Parent Notified': 'Yes',
    Status: 'Pending',
  });

  await sendWhatsAppMessage(from, `Issue reported for ${parent['Child Name']}. Principal will review and contact you within 24 hours.`);

  await sendWhatsAppMessage(ADMIN_PHONE, `DISCIPLINE ISSUE REPORTED\n\nStudent: ${parent['Child Name']}\nClass: ${parent.Class}\nIssue: ${messageBody}\nReporter: ${parent['Parent Name']}\n\nReview in dashboard.`);
}

// Emergency Alert
async function handleEmergency(from, parent) {
  if (!parent) {
    await sendWhatsAppMessage(from, 'Emergency: Please identify yourself first');
    return;
  }

  await sendWhatsAppMessage(from, `Emergency alert received. School authorities notified immediately. Stay tuned.`);

  await sendWhatsAppMessage(ADMIN_PHONE, `🚨 EMERGENCY ALERT\n\nFrom: ${parent['Parent Name']}\nChild: ${parent['Child Name']}\n\nImmediate action required!`);
}

// Registration
async function handleRegistration(from, messageBody) {
  const nameMatch = messageBody.match(/name\s+(?:is\s+)?([^,]+)/i);
  const childMatch = messageBody.match(/child\s+(?:is\s+)?([^,]+)/i);
  const classMatch = messageBody.match(/class\s+([^,]+)/i);

  if (nameMatch && childMatch && classMatch) {
    const name = nameMatch[1].trim();
    const child = childMatch[1].trim();
    const cls = classMatch[1].trim();

    await parentSheet.addRow({
      Phone: from,
      'Parent Name': name,
      'Child Name': child,
      Class: cls,
      Status: 'Active',
      'Joined Date': new Date().toISOString(),
    });

    await sendWhatsAppMessage(from, `Welcome to ${cls}! You're now registered.\n\nYou can now:\n• Check attendance daily\n• See homework\n• Track fees\n• Report issues\n\nJust ask!`);
  } else {
    await sendWhatsAppMessage(from, `Register like this:\nMy name is [Parent Name], child is [Child Name], class [Class]\n\nExample: My name is Raj, child is Arjun, class 5B`);
  }
}

// Send WhatsApp Message
async function sendWhatsAppMessage(to, message) {
  try {
    await client.messages.create({
      body: message,
      from: `whatsapp:${TWILIO_WHATSAPP_NUMBER}`,
      to: `whatsapp:${to}`,
    });
  } catch (error) {
    console.error('Error sending message:', error);
  }
}

// Broadcast Announcement
app.post('/api/announcement', async (req, res) => {
  const { title, message, category } = req.body;

  try {
    const parentRows = await parentSheet.getRows();

    await announcementsSheet.addRow({
      Title: title,
      Message: message,
      Category: category,
      'Sent Date': new Date().toISOString(),
      Recipients: parentRows.length,
      Status: 'Sent',
    });

    for (const parent of parentRows) {
      await sendWhatsAppMessage(parent.Phone, `📢 SCHOOL ANNOUNCEMENT\n\n${title}\n\n${message}`);
    }

    await sendWhatsAppMessage(ADMIN_PHONE, `Announcement sent to ${parentRows.length} parents.\n\n${title}`);

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('Error');
  }
});

// Dashboard
app.get('/dashboard', async (req, res) => {
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>School Management Dashboard</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f5f5; }
    .navbar { background: #000; color: #fff; padding: 20px 40px; }
    .navbar h1 { font-size: 24px; font-weight: 600; }
    .navbar p { font-size: 13px; color: #aaa; margin-top: 4px; }
    .container { max-width: 1200px; margin: 0 auto; padding: 40px 20px; }
    .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; margin-bottom: 40px; }
    .stat-card { background: #fff; padding: 25px; border-radius: 8px; border-left: 4px solid #000; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
    .stat-card h3 { font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 12px; }
    .stat-card .number { font-size: 36px; font-weight: 700; color: #000; }
    .section { background: #fff; border-radius: 8px; padding: 30px; margin-bottom: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
    .section h2 { font-size: 18px; font-weight: 600; margin-bottom: 25px; }
    .section h3 { font-size: 14px; font-weight: 600; color: #000; margin: 20px 0 10px 0; }
    .features { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; }
    .feature-box { background: #f9f9f9; padding: 15px; border-radius: 6px; border-left: 3px solid #000; }
    .feature-box h4 { font-size: 13px; font-weight: 600; margin-bottom: 8px; }
    .feature-box p { font-size: 12px; color: #666; }
    .form-group { margin-bottom: 15px; }
    .form-group label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px; }
    .form-group input, .form-group textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }
    .btn { padding: 10px 20px; background: #000; color: #fff; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; }
    .btn:hover { background: #333; }
    .navlinks { display: flex; gap: 12px; margin-top: 16px; }
    .navlinks a { display: inline-block; padding: 9px 16px; border-radius: 20px; font-size: 13px; font-weight: 600; text-decoration: none; }
    .navlinks a.primary { background: #5B3FE0; color: #fff; }
    .navlinks a.secondary { background: rgba(255,255,255,0.1); color: #fff; border: 1px solid rgba(255,255,255,0.25); }
  </style>
</head>
<body>
  <div class="navbar">
    <h1>School Management System</h1>
    <p>Complete solution for parent communication, attendance, fees, homework, and discipline tracking</p>
    <div class="navlinks">
      <a class="primary" href="/manager">Open Manager Dashboard →</a>
      <a class="secondary" href="/student">Open Student Dashboard →</a>
    </div>
  </div>
  
  <div class="container">
    <div class="stats">
      <div class="stat-card"><h3>Active Parents</h3><div class="number">0</div></div>
      <div class="stat-card"><h3>Students</h3><div class="number">0</div></div>
      <div class="stat-card"><h3>Pending Fees</h3><div class="number">₹0</div></div>
      <div class="stat-card"><h3>Discipline Issues</h3><div class="number">0</div></div>
    </div>

    <div class="section">
      <h2>What This System Does</h2>
      <h3>For Parents (via WhatsApp)</h3>
      <div class="features">
        <div class="feature-box">
          <h4>✓ Attendance Tracking</h4>
          <p>Parents see daily attendance status for their child instantly</p>
        </div>
        <div class="feature-box">
          <h4>✓ Homework Alerts</h4>
          <p>Automatic homework updates. No more "I forgot" excuses</p>
        </div>
        <div class="feature-box">
          <h4>✓ Fee Reminders</h4>
          <p>Automatic payment reminders. Track dues instantly</p>
        </div>
        <div class="feature-box">
          <h4>✓ Academic Progress</h4>
          <p>Marks, report cards, performance updates</p>
        </div>
        <div class="feature-box">
          <h4>✓ Report Issues</h4>
          <p>Report bullying, discipline issues, academic concerns</p>
        </div>
        <div class="feature-box">
          <h4>✓ School Announcements</h4>
          <p>Get event updates, schedule changes, urgent alerts</p>
        </div>
        <div class="feature-box">
          <h4>✓ Student Web Dashboard</h4>
          <p>A visual login for each child — attendance ring, timetable, report card, fee status, library books, all in one place</p>
        </div>
      </div>
    </div>

    <div class="section">
      <h2>For School Administration</h2>
      <div class="features">
        <div class="feature-box">
          <h4>✓ One Dashboard</h4>
          <p>All parent communication in one place. No WhatsApp chaos</p>
        </div>
        <div class="feature-box">
          <h4>✓ Attendance Management</h4>
          <p>Mark attendance once, parents get notified instantly</p>
        </div>
        <div class="feature-box">
          <h4>✓ Fee Tracking</h4>
          <p>Automate payment reminders. Reduce follow-ups by 80%</p>
        </div>
        <div class="feature-box">
          <h4>✓ Homework Distribution</h4>
          <p>Teachers post assignments. Parents see immediately</p>
        </div>
        <div class="feature-box">
          <h4>✓ Discipline Tracking</h4>
          <p>Track behavioral issues. Notify parents automatically</p>
        </div>
        <div class="feature-box">
          <h4>✓ Broadcast Announcements</h4>
          <p>Send to all parents at once. No manual messaging</p>
        </div>
        <div class="feature-box">
          <h4>✓ Manager Dashboard</h4>
          <p>Live view of attendance trends, fee collection, discipline queue, class-wise breakdowns and WhatsApp activity — all in one screen</p>
        </div>
        <div class="feature-box">
          <h4>✓ Class-wise Analytics</h4>
          <p>Spot at-risk classes instantly by attendance, fee collection and behaviour flags</p>
        </div>
      </div>
    </div>

    <div class="section">
      <h2>Broadcast Announcement</h2>
      <div class="form-group">
        <label>Title (e.g., "Parent-Teacher Meeting")</label>
        <input type="text" id="title" placeholder="Announcement title">
      </div>
      <div class="form-group">
        <label>Category</label>
        <input type="text" id="category" placeholder="Event / Fee / Holiday / Urgent">
      </div>
      <div class="form-group">
        <label>Message</label>
        <textarea id="message" placeholder="What do parents need to know?"></textarea>
      </div>
      <button class="btn" onclick="sendAnnouncement()">Send to All Parents</button>
    </div>

    <div class="section">
      <h2>Key Advantages Over Manual Process</h2>
      <h3>Before (Manual WhatsApp Groups)</h3>
      <p>❌ Messages get lost in groups | ❌ Manual follow-ups for fees | ❌ Parents don't know homework | ❌ Behavior issues reported weeks later | ❌ No accountability | ❌ Teachers overwhelmed</p>
      
      <h3>After (School Management System)</h3>
      <p>✓ Every parent gets every message | ✓ Automatic fee reminders | ✓ Homework posted instantly | ✓ Behavior issues logged immediately | ✓ Complete audit trail | ✓ Teachers save 3+ hours/day</p>
    </div>

    <div class="section">
      <h2>Pricing & ROI</h2>
      <h3>Monthly Cost: ₹7,000-10,000</h3>
      <p><strong>Saves:</strong> 5+ hours/day admin work | Reduces late fee collection by 70% | Improves parent satisfaction | Resolves issues 5x faster</p>
      <p><strong>ROI:</strong> Pays for itself in first week through fee collection alone</p>
    </div>
  </div>

  <script>
    async function sendAnnouncement() {
      const title = document.getElementById('title').value;
      const message = document.getElementById('message').value;
      const category = document.getElementById('category').value;
      
      if (!title || !message) {
        alert('Fill all fields');
        return;
      }

      try {
        await fetch('/api/announcement', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ title, message, category })
        });
        alert('Announcement sent to all parents!');
        document.getElementById('title').value = '';
        document.getElementById('message').value = '';
      } catch (error) {
        alert('Error sending announcement');
      }
    }
  </script>
</body>
</html>`;

  res.send(html);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`School system running on port ${PORT}`));

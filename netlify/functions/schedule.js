import { getStore } from '@netlify/blobs';

// This is the single source of truth for the schedule structure.
// It mirrors the uploaded Word document (Nexus_Individual_Meeting_Schedule_2026.docx).
// Note: the original document had a typo on Tuesday ("12;15") which is corrected here to "12:15".
function buildInitialData() {
  const days = [
    {
      day: 'Monday, Aug 24',
      slots: [
        { time: '8:00 – 9:00', name: null },
        { time: '9:00 – 10:00', name: null },
        { time: '10:00 – 11:00', name: null },
        { break: true, label: 'Break' },
        { time: '11:15 – 12:15', name: null },
        { time: '12:15 – 1:15', name: null },
        { break: true, label: 'Break' },
        { time: '1:30 – 2:30', name: null },
        { time: '2:30 – 3:30', name: null },
        { time: '3:30 – 4:30', name: null },
      ],
    },
    {
      day: 'Tuesday, Aug 25',
      slots: [
        { time: '10:15 – 11:15', name: null },
        { time: '11:15 – 12:15', name: null },
        { time: '12:15 – 1:15', name: null },
        { break: true, label: 'Break' },
        { time: '1:30 – 2:30', name: null },
        { time: '2:30 – 3:30', name: null },
        { time: '3:30 – 4:30', name: null },
        { break: true, label: 'Break' },
        { time: '5:00 – 6:00', name: null },
        { time: '6:00 – 7:00', name: null },
      ],
    },
    {
      day: 'Wednesday, Aug 26',
      slots: [
        { time: '10:15 – 11:15', name: null },
        { time: '11:15 – 12:15', name: null },
        { time: '12:15 – 1:15', name: null },
        { break: true, label: 'Break' },
        { time: '1:30 – 2:30', name: null },
        { time: '2:30 – 3:30', name: null },
        { time: '3:30 – 4:30', name: null },
        { break: true, label: 'Break' },
        { time: '5:00 – 6:00', name: null },
        { time: '6:00 – 7:00', name: null },
      ],
    },
    {
      day: 'Thursday, Aug 27',
      slots: [
        { time: '10:15 – 11:15', name: null },
        { time: '11:15 – 12:15', name: null },
        { time: '12:15 – 1:15', name: null },
        { break: true, label: 'Break' },
        { time: '1:30 – 2:30', name: null },
        { time: '2:30 – 3:30', name: null },
        { time: '3:30 – 4:30', name: null },
        { break: true, label: 'Break' },
        { time: '5:00 – 6:00', name: null },
        { time: '6:00 – 7:00', name: null },
      ],
    },
    {
      day: 'Friday, Aug 28',
      slots: [
        { time: '10:15 – 11:15', name: null },
        { time: '11:15 – 12:15', name: null },
        { time: '12:15 – 1:15', name: null },
        { break: true, label: 'Break' },
        { time: '1:30 – 2:30', name: null },
        { time: '2:30 – 3:30', name: null },
        { time: '3:30 – 4:30', name: null },
        { break: true, label: 'Break' },
        { time: '5:00 – 6:00', name: null },
        { time: '6:00 – 7:00', name: null },
      ],
    },
    {
      day: 'Monday, Aug 31',
      slots: [
        { time: '8:00 – 9:00', name: null },
        { time: '9:00 – 10:00', name: null },
        { time: '10:00 – 11:00', name: null },
        { break: true, label: 'Break' },
        { time: '11:15 – 12:15', name: null },
        { time: '12:15 – 1:15', name: null },
        { time: '1:15 – 2:15', name: null },
        { break: true, label: 'Break' },
        { time: '2:30 – 3:30', name: null },
        { time: '3:30 – 4:30', name: null },
      ],
    },
  ];
  return days;
}

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      'content-type': 'application/json',
      // Prevent any intermediate cache from serving a stale sign-up sheet.
      'cache-control': 'no-store',
    },
  });
}

export default async (req, context) => {
  const store = getStore('nexus-schedule');

  if (req.method === 'GET') {
    let data = await store.get('schedule', { type: 'json' });
    if (!data) {
      data = buildInitialData();
      await store.setJSON('schedule', data);
    }
    return jsonResponse(data);
  }

  if (req.method === 'POST') {
    let body;
    try {
      body = await req.json();
    } catch {
      return jsonResponse({ error: 'Invalid request body.' }, 400);
    }

    const { action, day, time, name, pin } = body || {};
    if (!action || !day || !time) {
      return jsonResponse({ error: 'Missing action, day, or time.' }, 400);
    }

    let data = await store.get('schedule', { type: 'json' });
    if (!data) data = buildInitialData();

    const dayEntry = data.find((d) => d.day === day);
    if (!dayEntry) return jsonResponse({ error: 'Unknown day.' }, 400);

    const slot = dayEntry.slots.find((s) => s.time === time);
    if (!slot || slot.break) return jsonResponse({ error: 'Unknown or non-bookable slot.' }, 400);

    if (action === 'signup') {
      const cleanName = String(name || '').trim().slice(0, 60);
      if (!cleanName) return jsonResponse({ error: 'Please enter a name.' }, 400);
      if (slot.name) {
        // Someone else took it between the person's page load and this request.
        return jsonResponse({ error: 'That slot was just taken by someone else. Please refresh and pick another.', data }, 409);
      }
      slot.name = cleanName;
    } else if (action === 'clear') {
      const adminPin = process.env.INSTRUCTOR_PIN || '1234';
      if (String(pin) !== String(adminPin)) {
        return jsonResponse({ error: 'Incorrect instructor PIN.' }, 403);
      }
      slot.name = null;
    } else {
      return jsonResponse({ error: 'Unknown action.' }, 400);
    }

    await store.setJSON('schedule', data);
    return jsonResponse(data);
  }

  return jsonResponse({ error: 'Method not allowed.' }, 405);
};

export const config = {
  path: '/api/schedule',
};

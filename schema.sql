CREATE TABLE IF NOT EXISTS tickets (
    ticket_id TEXT PRIMARY KEY, 
    title TEXT, 
    status TEXT, 
    priority TEXT,
    request_type TEXT, 
    staff_assigned TEXT, 
    requester TEXT, 
    date_opened DATE, 
    location TEXT, 
    description TEXT, 
    resolution_notes TEXT
);

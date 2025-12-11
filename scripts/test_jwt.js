const tokenPayload = "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImlzTGc0MFNoZDYyTW1RdThKYk5NdXdUYV9VZzFNREVLbTdfUnhxeTd6OHMiLCJhbGciOiJSUzI1NiIsIng1dCI6InJ0c0ZULWItN0x1WTdEVlllU05LY0lKN1ZuYyIsImtpZCI6InJ0c0ZULWItN0x1WTdEVlllU05LY0lKN1ZuYyJ9.eyJhdWQiOiJodHRwczovL2ljMy50ZWFtcy5vZmZpY2UuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvMzFmNjYwYTUtMTkyYS00ZGIzLTkyYmEtY2E0MjRmMWIyNTllLyIsImlhdCI6MTc2MzkzNTQyNiwibmJmIjoxNzYzOTM1NDI2LCJleHAiOjE3NjM5Mzk1NTMsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBWFFBaS84YUFBQUFIdjN3TDhHcGJteE4rclowcHRPRXc2ZlJhR1pqQmFTbE5KUFFKd1hqMkwrN1JJRHc0dHlFMVZoeENzYzRNOE5ZVUlUbGR5NEQ1akZ1K3hlUUpSc1J1ZGk2UThOcVF4LzdGTzNUWENUR3lnVzE2N0w5V1p3MUVqYzB1NWIzR0g5bFlQRVlOMDdZdFdCdGJFK0NZa0JpM2c9PSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwaWQiOiI1ZTNjZTZjMC0yYjFmLTQyODUtOGQ0Yi03NWVlNzg3ODczNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkphbGlsIiwiZ2l2ZW5fbmFtZSI6IkFobWFkIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMjA3LjgxLjQuMTMyIiwibmFtZSI6IkphbGlsLCBBaG1hZCBbTkhdIiwib2lkIjoiNDZkZTFhMjctMzEwNi00NzhiLWJkNDktNmY2NzVmODg4NDhkIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTEwNzA3NTkxNTEtMTE0NDMzODM0Ny0xOTA1MjAzODg1LTE5MDk2NCIsInB1aWQiOiIxMDAzMjAwMzdGRDIxRThBIiwicmgiOiIxLkFXNEFwV0QyTVNvWnMwMlN1c3BDVHhzbG5sVHdxam1sZ2NkSXBQZ0Nrd0VnbGJsZUFlSnVBQS4iLCJzY3AiOiJUZWFtcy5BY2Nlc3NBc1VzZXIuQWxsIiwic2lkIjoiMDBhYjYzMTktYmQ3Mi0zNGY5LTJmM2UtMDE2ZTViMTA1ODliIiwic3ViIjoiWHFHZUk3akppaHdmMjA4a19UNEFoX1k4Ym1NdTRvYzNzODNQc3ItVjBqbyIsInRpZCI6IjMxZjY2MGE1LTE5MmEtNGRiMy05MmJhLWNhNDI0ZjFiMjU5ZSIsInVuaXF1ZV9uYW1lIjoiQWhtYWQuSmFsaWxAbm9ydGhlcm5oZWFsdGguY2EiLCJ1cG4iOiJBaG1hZC5KYWxpbEBub3J0aGVybmhlYWx0aC5jYSIsInV0aSI6Im5VTlB6NW93UTBpNVphQTdmU2NjQUEiLCJ2ZXIiOiIxLjAiLCJ4bXNfYWN0X2ZjdCI6IjUgMyIsInhtc19jYyI6WyJDUDEiXSwieG1zX2Z0ZCI6Iml1SlV4dnNEbVlPNTZXNWVpdmFScWpvOGhqXzkycF9iZDJHQ2pJV0E0cmtCZFhOemIzVjBhQzFrYzIxeiIsInhtc19pZHJlbCI6IjEgMjAiLCJ4bXNfc3NtIjoiMSIsInhtc19zdWJfZmN0IjoiMyAyIn0.hOSzxLtzLzS2WcuRWpK3ffYoRx8gN8TmQ2155rEEEGV0PLkeuQES-25Wk0xdISVxH_pvZKLZ7SJ_I9eDjTiPPklzeIak7FglAiv6C6Vy9nUm-9P1vZcK-D6Kaes-JIFyyuysnmFKQA-ePkoBOKSQ6I4VEtgTbgsLvLzWzuLCxNGPVc5ED8vN76_LJNqvXvAzIEmDZlvIat7vej38d3RY9BWzRG54anA9FDYWhJLWGKE5-qFY2WXyjq-kJE69thQdmvH66cVw8RrE4Viw_33N-6BHKy7qGMJc2opD1tsdIsNnkIIKF7miYgL3XJvqg50NCKOndC1vpNPATRSaH3D29A";

function isValidTeamsToken(token) {
    try {
        // Simulate atob in node
        let base64 = token.replace(/-/g, '+').replace(/_/g, '/');
        const pad = base64.length % 4;
        if (pad) {
            base64 += new Array(5 - pad).join('=');
        }
        const decoded = Buffer.from(base64, 'base64').toString('utf-8');
        console.log('Decoded:', decoded);
        const payload = JSON.parse(decoded);
        console.log('Audience:', payload.aud);
        if (payload.aud === 'https://ic3.teams.office.com') return true;
        return false;
    } catch (e) {
        console.error('Decoding failed:', e);
        return false;
    }
}

console.log('Is Valid:', isValidTeamsToken(tokenPayload));

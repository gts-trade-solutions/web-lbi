# MySQL + S3 Migration Runbook

## 1. Create MySQL database
```sql
CREATE DATABASE tracker;
```

## 2. Import schema
```bash
mysql -u root -p tracker < database/schema.sql
```

## 3. Configure `.env.local`
Use the variables from `.env.example`:
- `MYSQL_HOST`
- `MYSQL_PORT`
- `MYSQL_USER`
- `MYSQL_PASSWORD`
- `MYSQL_DATABASE`
- `AWS_S3_REGION`
- `AWS_S3_ACCESS_KEY_ID`
- `AWS_S3_SECRET_ACCESS_KEY`
- `AWS_S3_BUCKET_NAME`
- `NEXT_PUBLIC_S3_BUCKET_URL`
- `NEXT_PUBLIC_API_BASE_URL`
- `EXPO_PUBLIC_API_BASE_URL`
- `JWT_SECRET_KEY`
- `AUTH_TOKEN_EXPIRES_IN`

## 4. Install packages
```bash
npm install
```

## 4.1 Seed/refresh admin user (local)
```bash
node scripts/create-admin-user.js
```

## 5. Run backend
```bash
npm run dev
```

## 6. Test health route
Open:
`http://localhost:3000/api/health`

Expected:
```json
{
  "ok": true,
  "database": "connected"
}
```

## 7. Mobile app notes
- Set `EXPO_PUBLIC_API_BASE_URL` to your PC LAN IP URL, e.g. `http://192.168.x.x:3000`.
- Keep web using `NEXT_PUBLIC_API_BASE_URL`.

## 8. Production VPS notes
- Use domain-based URL for `NEXT_PUBLIC_API_BASE_URL`/`EXPO_PUBLIC_API_BASE_URL`, not LAN IP.
- Run Next.js with PM2:
  - `npm run build`
  - `pm2 start npm --name lbi-web -- start`
- Configure Nginx reverse proxy to Next.js app port.
- Keep MySQL and AWS secrets only on the server environment.

const { Pool } = require('pg');

async function resetDatabase() {
  const pool = new Pool({
    connectionString: 'postgres://postgres:postgres@localhost:5432/ops_db'
  });

  const client = await pool.connect();
  try {
    console.log('Dropping public schema...');
    await client.query('DROP SCHEMA public CASCADE');
    
    console.log('Creating public schema...');
    await client.query('CREATE SCHEMA public');
    
    console.log('Database reset complete!');
  } catch (error) {
    console.error('Error:', error.message);
  } finally {
    client.release();
    await pool.end();
  }
}

resetDatabase();

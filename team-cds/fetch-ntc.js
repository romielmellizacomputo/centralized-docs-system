fetch('https://script.google.com/macros/s/AKfycbxby6scKsSziC5LBr2kFgd5_0UCx56X6K8qyuDbZNYxJBu_nhln2oDYQBiFpDAOHrmjyQ/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));

### ProductRequest (REST)

- `GET /api/product-requests`
    - Query params: `number`, `type`, `processStatus`, `requestedById`, `issuedById`, `workOrderId`
- `GET /api/product-requests/{id}`
- `POST /api/product-requests?autoNumber={true|false}` (create)
- `PUT /api/product-requests/{id}` (update)
- `DELETE /api/product-requests/{id}`

### ProductRequestItem (REST)

- `GET /api/product-request-items`
    - Query params: `productRequestId`, `productRequestIds` (comma-separated), `productId`, `productLotId`, `equipmentId`
- `GET /api/product-request-items/{id}`
- `POST /api/product-request-items` (create)
- `POST /api/product-request-items/bulk` (upsert list)
- `PUT /api/product-request-items/{id}` (update)
- `DELETE /api/product-request-items/{id}`


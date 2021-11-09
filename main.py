from http.server import HTTPServer, SimpleHTTPRequestHandler


server = HTTPServer(('0.0.0.0', 8001), SimpleHTTPRequestHandler)
server.serve_forever()

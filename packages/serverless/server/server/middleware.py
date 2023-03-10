from django.http import HttpRequest, HttpResponse
from django.utils.deprecation import MiddlewareMixin


class CorsMiddleware(MiddlewareMixin):
    headers = [
        "accept",
        "accept-encoding",
        "authorization",
        "content-type",
        "dnt",
        "origin",
        "user-agent",
        "x-csrftoken",
        "x-requested-with",
    ]

    methods = ["DELETE", "GET", "OPTIONS", "PATCH", "POST", "PUT"]

    def process_request(self, request: HttpRequest) -> HttpResponse | None:
        if request.method == "OPTIONS" and "HTTP_ACCESS_CONTROL_REQUEST_METHOD" in request.META:
            response = HttpResponse()
            response["Content-Length"] = "0"
            return response
        return None

    def process_response(self, request: HttpRequest, response: HttpResponse) -> HttpResponse:
        origin = request.META.get("HTTP_ORIGIN")

        if not origin:
            return response

        response["Access-Control-Allow-Credentials"] = "true"

        # TODO: We should check if `origin` matches one of the license domains
        response["Access-Control-Allow-Origin"] = origin

        if request.method == "OPTIONS":
            response["Access-Control-Allow-Headers"] = ", ".join(self.headers)
            response["Access-Control-Allow-Methods"] = ", ".join(self.methods)

        return response

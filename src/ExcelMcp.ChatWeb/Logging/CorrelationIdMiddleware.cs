using Serilog.Context;

namespace ExcelMcp.ChatWeb.Logging;

/// <summary>
/// Middleware to add correlation ID to HTTP request logging context.
/// </summary>
public class CorrelationIdMiddleware
{
    private readonly RequestDelegate _next;
    private const string CorrelationIdHeaderName = "X-Correlation-Id";

    public CorrelationIdMiddleware(RequestDelegate next)
    {
        _next = next;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        // Get or generate correlation ID
        string correlationId = context.Request.Headers[CorrelationIdHeaderName].FirstOrDefault()
                              ?? Guid.NewGuid().ToString();

        // Add to response headers for client tracking
        context.Response.Headers[CorrelationIdHeaderName] = correlationId;

        // Push to logging context
        using (LogContext.PushProperty("CorrelationId", correlationId))
        {
            // Store in HttpContext for access by services
            context.Items["CorrelationId"] = correlationId;

            await _next(context);
        }
    }
}

/// <summary>
/// Extension method for registering correlation ID middleware.
/// </summary>
public static class CorrelationIdMiddlewareExtensions
{
    public static IApplicationBuilder UseCorrelationId(this IApplicationBuilder builder)
    {
        return builder.UseMiddleware<CorrelationIdMiddleware>();
    }
}

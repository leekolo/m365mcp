using FluentAssertions;
using NewMCP.Tools;
using Xunit;

namespace NewMCP.Tests;

public class RoleBasedAccessControlTests
{
    [Fact]
    public void ValidateAccess_WithDefaultRole_AllowsSharePointAccess()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var (hasAccess, errorMessage) = accessControl.ValidateAccess("user1", "SharePoint");

        // Assert
        hasAccess.Should().BeTrue();
        errorMessage.Should().BeEmpty();
    }

    [Fact]
    public void ValidateAccess_WithDefaultRole_DeniesEmailAccess()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var (hasAccess, errorMessage) = accessControl.ValidateAccess("user1", "Email");

        // Assert
        hasAccess.Should().BeFalse();
        errorMessage.Should().Contain("Email");
    }

    [Fact]
    public void GetUserRole_DefaultUser_ReturnsSharePointOnly()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var role = accessControl.GetUserRole("user1");

        // Assert
        role.Should().Be("SharePoint.Only");
    }

    [Fact]
    public void HasAccess_SharePoint_ReturnsTrue()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var hasAccess = accessControl.HasAccess("user1", "SharePoint");

        // Assert
        hasAccess.Should().BeTrue();
    }

    [Fact]
    public void HasAccess_Email_ReturnsFalse()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var hasAccess = accessControl.HasAccess("user1", "Email");

        // Assert
        hasAccess.Should().BeFalse();
    }

    [Fact]
    public void GetAllowedTools_ReturnsSharePointOnly()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var tools = accessControl.GetAllowedTools("user1");

        // Assert
        tools.Should().Contain("SharePoint");
        tools.Should().NotContain("Email");
    }

    [Fact]
    public void GetRoleAssignments_NoMapping_ReturnsDefaultConfig()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var result = accessControl.GetRoleAssignments();

        // Assert
        result.Should().Contain("configured");
        result.Should().Contain("false");
    }

    [Fact]
    public void ValidateAccess_UnknownUser_UsesDefaultRole()
    {
        // Arrange
        var accessControl = new RoleBasedAccessControl();

        // Act
        var (sharePointAccess, _) = accessControl.ValidateAccess("unknown-user", "SharePoint");
        var (emailAccess, _) = accessControl.ValidateAccess("unknown-user", "Email");

        // Assert
        sharePointAccess.Should().BeTrue();
        emailAccess.Should().BeFalse();
    }
}
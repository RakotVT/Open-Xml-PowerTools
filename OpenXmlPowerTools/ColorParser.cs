// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;
using SkiaSharp;

namespace OpenXmlPowerTools;

public static class ColorParser
{
    public static SKColor FromName(string name)
    {
        var skColor = typeof(SKColors).GetField(name); 
        if (skColor != null) 
            return (SKColor)skColor.GetValue(null)!; 
        return SKColors.Empty;
    }

    public static bool TryFromName(string name, out SKColor color)
    {
        try
        {
            color = FromName(name);
            return color != default;
        }
        catch
        {
            color = default;
            return false;
        }
    }

    public static bool IsValidName(string name)
    {
        return TryFromName(name, out _);
    }
}

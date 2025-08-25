"""
Windows Volume Control MCP Server

This MCP server provides tools to control Windows system volume including:
- Get current volume level
- Set volume to specific percentage
- Mute/unmute system
- Toggle mute status
"""

import logging
from typing import Any

from mcp.server.fastmcp import FastMCP

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("volume-control-server")

# Create the MCP server
mcp = FastMCP("Volume Control")

class VolumeController:
    """Handles all volume control operations using pycaw library"""
    
    def __init__(self):
        self._volume = None
        self._initialize_audio()
    
    def _initialize_audio(self):
        """Initialize the audio interface"""
        try:
            from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
            from comtypes import CLSCTX_ALL
            
            # Get the default audio device
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self._volume = interface.QueryInterface(IAudioEndpointVolume)
            
            logger.info("Audio interface initialized successfully")
            
        except ImportError as e:
            logger.error(f"Required libraries not installed: {e}")
            raise
        except Exception as e:
            logger.error(f"Failed to initialize audio interface: {e}")
            raise
    
    def get_volume(self) -> dict:
        """Get current volume level and mute status"""
        try:
            if not self._volume:
                raise RuntimeError("Audio interface not initialized")
            
            # Get current volume information
            volume_scalar = self._volume.GetMasterVolumeLevelScalar()  # 0.0 to 1.0
            volume_db = self._volume.GetMasterVolumeLevel()           # Decibels
            is_muted = self._volume.GetMute()                         # Boolean
            
            # Convert to percentage
            volume_percentage = round(volume_scalar * 100, 1)
            
            return {
                "volume_percentage": volume_percentage,
                "volume_scalar": volume_scalar,
                "volume_db": round(volume_db, 2),
                "is_muted": bool(is_muted),
                "status": "success"
            }
            
        except Exception as e:
            logger.error(f"Failed to get volume: {e}")
            return {
                "error": str(e),
                "status": "error"
            }
    
    def set_volume(self, percentage: float) -> dict:
        """Set volume to specific percentage (0-100)"""
        try:
            if not self._volume:
                raise RuntimeError("Audio interface not initialized")
            
            # Validate input
            if not 0 <= percentage <= 100:
                raise ValueError("Volume percentage must be between 0 and 100")
            
            # Convert percentage to scalar (0.0 to 1.0)
            volume_scalar = percentage / 100.0
            
            # Set the volume
            self._volume.SetMasterVolumeLevelScalar(volume_scalar, None)
            
            # Get the actual set volume (might be slightly different due to hardware)
            actual_volume = self._volume.GetMasterVolumeLevelScalar()
            actual_percentage = round(actual_volume * 100, 1)
            
            logger.info(f"Volume set to {actual_percentage}%")
            
            return {
                "requested_percentage": percentage,
                "actual_percentage": actual_percentage,
                "volume_scalar": actual_volume,
                "status": "success"
            }
            
        except Exception as e:
            logger.error(f"Failed to set volume: {e}")
            return {
                "error": str(e),
                "status": "error"
            }
    
    def mute(self) -> dict:
        """Mute the system"""
        try:
            if not self._volume:
                raise RuntimeError("Audio interface not initialized")
            
            # Check current mute status
            was_muted = self._volume.GetMute()
            
            if was_muted:
                return {
                    "message": "System was already muted",
                    "was_muted": True,
                    "is_muted": True,
                    "status": "success"
                }
            
            # Mute the system
            self._volume.SetMute(1, None)
            
            logger.info("System muted")
            
            return {
                "message": "System muted successfully",
                "was_muted": False,
                "is_muted": True,
                "status": "success"
            }
            
        except Exception as e:
            logger.error(f"Failed to mute: {e}")
            return {
                "error": str(e),
                "status": "error"
            }
    
    def unmute(self) -> dict:
        """Unmute the system"""
        try:
            if not self._volume:
                raise RuntimeError("Audio interface not initialized")
            
            # Check current mute status
            was_muted = self._volume.GetMute()
            
            if not was_muted:
                return {
                    "message": "System was already unmuted",
                    "was_muted": False,
                    "is_muted": False,
                    "status": "success"
                }
            
            # Unmute the system
            self._volume.SetMute(0, None)
            
            logger.info("System unmuted")
            
            return {
                "message": "System unmuted successfully",
                "was_muted": True,
                "is_muted": False,
                "status": "success"
            }
            
        except Exception as e:
            logger.error(f"Failed to unmute: {e}")
            return {
                "error": str(e),
                "status": "error"
            }
    
    def toggle_mute(self) -> dict:
        """Toggle mute status"""
        try:
            if not self._volume:
                raise RuntimeError("Audio interface not initialized")
            
            # Get current mute status
            current_mute = self._volume.GetMute()
            
            # Toggle mute
            new_mute = not current_mute
            self._volume.SetMute(1 if new_mute else 0, None)
            
            action = "muted" if new_mute else "unmuted"
            logger.info(f"System {action} (toggled)")
            
            return {
                "message": f"System {action} successfully",
                "was_muted": bool(current_mute),
                "is_muted": bool(new_mute),
                "action": action,
                "status": "success"
            }
            
        except Exception as e:
            logger.error(f"Failed to toggle mute: {e}")
            return {
                "error": str(e),
                "status": "error"
            }


# Initialize the volume controller
volume_controller = VolumeController()

@mcp.tool()
def get_volume() -> str:
    """
    Get the current system volume level and mute status.
    Returns volume percentage, mute status, and decibel level.
    """
    result = volume_controller.get_volume()
    
    if result["status"] == "success":
        return f"🔊 Volume: {result['volume_percentage']}% | Muted: {'Yes' if result['is_muted'] else 'No'} | dB: {result['volume_db']}"
    else:
        return f"❌ Error: {result['error']}"

@mcp.tool()
def set_volume(percentage: float) -> str:
    """
    Set the system volume to a specific percentage.
    
    Args:
        percentage: Volume percentage (0-100)
    """
    result = volume_controller.set_volume(percentage)
    
    if result["status"] == "success":
        return f"🔊 Volume set to {result['actual_percentage']}% (requested: {result['requested_percentage']}%)"
    else:
        return f"❌ Error: {result['error']}"

@mcp.tool()
def mute() -> str:
    """
    Mute the system audio.
    """
    result = volume_controller.mute()
    
    if result["status"] == "success":
        return f"� {result['message']}"
    else:
        return f"❌ Error: {result['error']}"

@mcp.tool()
def unmute() -> str:
    """
    Unmute the system audio.
    """
    result = volume_controller.unmute()
    
    if result["status"] == "success":
        return f"🔇 {result['message']}"
    else:
        return f"❌ Error: {result['error']}"

@mcp.tool()
def toggle_mute() -> str:
    """
    Toggle the system mute status (mute if unmuted, unmute if muted).
    """
    result = volume_controller.toggle_mute()
    
    if result["status"] == "success":
        return f"🔇 {result['message']}"
    else:
        return f"❌ Error: {result['error']}"

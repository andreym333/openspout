<?php

namespace OpenSpout\Writer\Common\Manager\Style;

use OpenSpout\Common\Entity\Style\Style;

/**
 * Registry for all used styles.
 */
class StyleRegistry
{
    /** @var array [SERIALIZED_STYLE] => [STYLE_ID] mapping table, keeping track of the registered styles */
    protected $serializedStyleToStyleIdMappingTable = [];

    /** @var array [STYLE_ID] => [STYLE] mapping table, keeping track of the registered styles */
    protected $styleIdToStyleMappingTable = [];

    public function __construct(Style $defaultStyle)
    {
        // This ensures that the default style is the first one to be registered
        $this->registerStyle($defaultStyle);
    }

    /**
     * Registers the given style as a used style.
     * Duplicate styles won't be registered more than once.
     *
     * @param Style $style The style to be registered
     *
     * @return Style the registered style, updated with an internal ID
     */
    public function registerStyle(Style $style)
    {
        $serializedStyle = $style->serialize();

        if (!isset($this->serializedStyleToStyleIdMappingTable[$serializedStyle])) {
            $nextStyleId = \count($this->serializedStyleToStyleIdMappingTable);
            $style->markAsRegistered($nextStyleId);

            $this->serializedStyleToStyleIdMappingTable[$serializedStyle] = $nextStyleId;
            $this->styleIdToStyleMappingTable[$nextStyleId] = $style;
            return $style;
        }

        $styleId = $this->serializedStyleToStyleIdMappingTable[$serializedStyle];
        return $this->styleIdToStyleMappingTable[$styleId];
    }

    /**
     * @return Style[] List of registered styles
     */
    public function getRegisteredStyles()
    {
        return array_values($this->styleIdToStyleMappingTable);
    }

    /**
     * @param int $styleId
     *
     * @return Style
     */
    public function getStyleFromStyleId($styleId)
    {
        return $this->styleIdToStyleMappingTable[$styleId];
    }
}
